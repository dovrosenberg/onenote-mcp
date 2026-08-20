#!/usr/bin/env bash
#
# Provisions every GCP resource the onenote-mcp deploy pipeline needs:
# APIs, an Artifact Registry Docker repo, a Firestore database in Native mode,
# a runtime and a deploy service account, and a Workload Identity Federation
# pool bound to this GitHub repository only.
#
# The script is idempotent -- re-running it on an already-configured project
# is a no-op that still re-asserts every IAM binding and prints the output
# block at the end.
#
# No service-account JSON key is created. Deploys authenticate through
# Workload Identity Federation, so no key is needed anywhere.
#
# Issue #3 is the manual runbook for executing this script.

set -euo pipefail

# ---------------------------------------------------------------------------
# Parameters. Override any of these from the environment, or edit in place.
#
# PROJECT defaults to the project this repository deploys to. Set PROJECT in
# the environment to provision a different one; the script creates resources
# in whichever project it resolves, so check the "Project:" line it prints
# before letting it run.
#
# REGION is used both for Cloud Run and for the Firestore database, so it
# must be a valid Firestore location name (e.g. us-central1, europe-west1,
# nam5). Firestore locations are a subset of Cloud Run regions.
# ---------------------------------------------------------------------------
PROJECT="${PROJECT:-onenote-mcp-505918}"
REGION="${REGION:-us-central1}"
GAR_REGION="${GAR_REGION:-us-central1}"
SERVICE="${SERVICE:-onenote-mcp}"
GITHUB_REPO="${GITHUB_REPO:-dovrosenberg/onenote-mcp}"

# Derived names. Not intended to be overridden.
RUNTIME_SA_ID="${SERVICE}-run"
DEPLOY_SA_ID="${SERVICE}-deploy"
GAR_REPO="${SERVICE}"
POOL_ID="github-pool"
PROVIDER_ID="github-provider"

# The bucket holding rendered ink PNGs and any page HTML too large for a
# Firestore document (issue #30). Bucket names are globally unique across all
# of GCS, so the project number is appended to make a collision with someone
# else's bucket impossible. Override MIRROR_BUCKET to use a name you already
# own.
MIRROR_BUCKET="${MIRROR_BUCKET:-${SERVICE}-mirror-${PROJECT##*-}}"

log()  { printf '==> %s\n' "$*"; }
warn() { printf 'WARNING: %s\n' "$*" >&2; }
die()  { printf 'ERROR: %s\n' "$*" >&2; exit 1; }

# ---------------------------------------------------------------------------
# Preflight
# ---------------------------------------------------------------------------
command -v gcloud >/dev/null 2>&1 || die "gcloud not found on PATH"
[[ -n "$PROJECT" ]] || die "Set PROJECT (env var, or edit the top of this script)"

gcloud projects describe "$PROJECT" >/dev/null 2>&1 \
  || die "Cannot describe project $PROJECT. Run: gcloud auth login"

RUNTIME_SA="${RUNTIME_SA_ID}@${PROJECT}.iam.gserviceaccount.com"
DEPLOY_SA="${DEPLOY_SA_ID}@${PROJECT}.iam.gserviceaccount.com"

# Non-fatal: checking billing needs a permission the operator may not have
# granted themselves, and a false negative should not block provisioning.
billing_enabled="$(gcloud beta billing projects describe "$PROJECT" \
  --format='value(billingEnabled)' 2>/dev/null || true)"
if [[ -z "$billing_enabled" ]]; then
  warn "Could not check billing on $PROJECT. Confirm it is enabled before deploying."
elif [[ "$billing_enabled" != "True" ]]; then
  warn "Billing is not enabled on $PROJECT. Artifact Registry and Cloud Run will fail."
fi

log "Project:      $PROJECT"
log "Region:       $REGION"
log "GAR region:   $GAR_REGION"
log "Service:      $SERVICE"
log "GitHub repo:  $GITHUB_REPO"
log "Runtime SA:   $RUNTIME_SA"
log "Deploy SA:    $DEPLOY_SA"

# ---------------------------------------------------------------------------
# ensure_resource <human name> <describe command...> -- <create command...>
#
# Runs the describe command; creates only if it fails. Every create command
# in this script exits non-zero with ALREADY_EXISTS on a second run, so the
# guard is what makes the script re-runnable.
# ---------------------------------------------------------------------------
ensure_resource() {
  local name="$1"; shift
  local -a describe_cmd=() create_cmd=()
  local seen_sep=0 arg
  for arg in "$@"; do
    if [[ "$arg" == "--" && $seen_sep -eq 0 ]]; then
      seen_sep=1
      continue
    fi
    if [[ $seen_sep -eq 0 ]]; then
      describe_cmd+=("$arg")
    else
      create_cmd+=("$arg")
    fi
  done
  if "${describe_cmd[@]}" >/dev/null 2>&1; then
    log "$name already exists, skipping create"
  else
    log "creating $name"
    "${create_cmd[@]}"
  fi
}

# ---------------------------------------------------------------------------
# APIs. `services enable` is idempotent, so this needs no guard. One call so
# the operator waits once rather than six times.
# ---------------------------------------------------------------------------
log "enabling APIs"
# cloudscheduler is enabled here although the keepalive job is optional and this
# script does not create it: the job needs MCP_KEEPALIVE_SECRET, which is generated
# by the operator and unknown here. Enabling an API costs nothing, and the README's
# keepalive command fails with a PERMISSION_DENIED that reads like a missing role
# when the API is off.
gcloud services enable \
  run.googleapis.com \
  artifactregistry.googleapis.com \
  firestore.googleapis.com \
  storage.googleapis.com \
  iamcredentials.googleapis.com \
  sts.googleapis.com \
  cloudresourcemanager.googleapis.com \
  cloudscheduler.googleapis.com \
  --project="$PROJECT"

# ---------------------------------------------------------------------------
# Artifact Registry
# ---------------------------------------------------------------------------
ensure_resource "Artifact Registry repo $GAR_REPO ($GAR_REGION)" \
  gcloud artifacts repositories describe "$GAR_REPO" \
    --location="$GAR_REGION" --project="$PROJECT" \
  -- \
  gcloud artifacts repositories create "$GAR_REPO" \
    --repository-format=docker \
    --location="$GAR_REGION" \
    --project="$PROJECT" \
    --description="Container images for $SERVICE"

# ---------------------------------------------------------------------------
# Firestore, Native mode.
#
# --location must be a Firestore location name, not any Cloud Run region.
# The database's location cannot be changed after creation; moving it means
# deleting and recreating the database.
# ---------------------------------------------------------------------------
ensure_resource "Firestore Native database ($REGION)" \
  gcloud firestore databases describe --database='(default)' --project="$PROJECT" \
  -- \
  gcloud firestore databases create \
    --location="$REGION" \
    --type=firestore-native \
    --project="$PROJECT"

# ---------------------------------------------------------------------------
# Cloud Storage bucket for the page mirror (issue #30).
#
# --uniform-bucket-level-access turns off per-object ACLs, so access is decided
# by IAM alone and an object cannot be made public one at a time.
# --public-access-prevention refuses any binding that would make it public at
# all. Both matter here rather than being boilerplate: the objects are rendered
# handwriting, which CLAUDE.md's hygiene rules name as a thing that must never
# leave this account.
#
# Same location as the service so a read is in-region. Nothing here sets a
# lifecycle rule: the sync deletes an object when it deletes the page it
# belongs to, and a rule that expired objects by age would delete the ink of a
# page nobody has edited in a year.
# ---------------------------------------------------------------------------
ensure_resource "Cloud Storage bucket $MIRROR_BUCKET ($REGION)" \
  gcloud storage buckets describe "gs://$MIRROR_BUCKET" --project="$PROJECT" \
  -- \
  gcloud storage buckets create "gs://$MIRROR_BUCKET" \
    --location="$REGION" \
    --uniform-bucket-level-access \
    --public-access-prevention \
    --project="$PROJECT"

# ---------------------------------------------------------------------------
# Service accounts
# ---------------------------------------------------------------------------
ensure_resource "runtime service account $RUNTIME_SA" \
  gcloud iam service-accounts describe "$RUNTIME_SA" --project="$PROJECT" \
  -- \
  gcloud iam service-accounts create "$RUNTIME_SA_ID" \
    --display-name="$SERVICE Cloud Run runtime" \
    --project="$PROJECT"

ensure_resource "deploy service account $DEPLOY_SA" \
  gcloud iam service-accounts describe "$DEPLOY_SA" --project="$PROJECT" \
  -- \
  gcloud iam service-accounts create "$DEPLOY_SA_ID" \
    --display-name="$SERVICE GitHub Actions deployer" \
    --project="$PROJECT"

# ---------------------------------------------------------------------------
# Project-level role grants.
#
# add-iam-policy-binding is idempotent, so these are re-asserted on every run
# rather than guarded. --condition=None suppresses the interactive prompt that
# gcloud raises when the existing policy contains conditional bindings; without
# it a re-run blocks waiting for input.
# ---------------------------------------------------------------------------
log "granting roles/datastore.user to $RUNTIME_SA"
gcloud projects add-iam-policy-binding "$PROJECT" \
  --member="serviceAccount:$RUNTIME_SA" \
  --role="roles/datastore.user" \
  --condition=None --quiet >/dev/null

# Scoped to the bucket, not the project. The runtime service account creates,
# reads and deletes the mirror's ink and HTML objects; a project-level grant
# would reach every bucket in the project instead of only this one. Same
# reasoning as the serviceAccountUser binding further down.
log "granting roles/storage.objectAdmin on gs://$MIRROR_BUCKET to $RUNTIME_SA"
gcloud storage buckets add-iam-policy-binding "gs://$MIRROR_BUCKET" \
  --member="serviceAccount:$RUNTIME_SA" \
  --role="roles/storage.objectAdmin" \
  --project="$PROJECT" --quiet >/dev/null

for role in roles/run.admin roles/artifactregistry.writer; do
  log "granting $role to $DEPLOY_SA"
  gcloud projects add-iam-policy-binding "$PROJECT" \
    --member="serviceAccount:$DEPLOY_SA" \
    --role="$role" \
    --condition=None --quiet >/dev/null
done

# Scoped to the runtime service account resource, not the project. The deploy
# job needs to act as the runtime SA when it passes --service-account to the
# Cloud Run deploy; granting this at the project level would let the deploy SA
# impersonate every service account in the project instead of just this one.
log "granting roles/iam.serviceAccountUser on $RUNTIME_SA to $DEPLOY_SA"
gcloud iam service-accounts add-iam-policy-binding "$RUNTIME_SA" \
  --member="serviceAccount:$DEPLOY_SA" \
  --role="roles/iam.serviceAccountUser" \
  --project="$PROJECT" --quiet >/dev/null

# ---------------------------------------------------------------------------
# Firestore composite indexes for the page mirror (issue #30).
#
# Firestore indexes single fields automatically and composite queries not at
# all. A query with no index fails at runtime with FAILED_PRECONDITION and a
# console link -- which is a fine developer experience and a bad production
# one, because nothing in CI notices and the first person to hit it is a model
# mid-conversation. So they are created here.
#
# `indexes composite create` has no describe-by-definition form to guard on, so
# ensure_resource does not fit: a second run answers ALREADY_EXISTS, which is
# the idempotent outcome and not a failure. The `|| true` is that, and only
# that -- a real failure still prints its own error above it.
#
# --async is not optional here. Without it gcloud submits the operation and
# then blocks polling until the index has finished BUILDING, which takes
# minutes per index and looks exactly like a hung script -- it did, on the
# first real run. The build continues server-side either way; the state is
# visible in `gcloud firestore indexes composite list`, and the next section
# of this script prints the command. An index that is still CREATING when the
# service queries it fails that query with FAILED_PRECONDITION, so wait for
# READY before turning the mirror on.
#
# Collection ids are the last segment of MIRROR_ROOT_DOC's subcollections, so
# they are fixed strings rather than derived: a collection-group index is
# keyed by the collection id alone, and every deployment uses the same three.
# ---------------------------------------------------------------------------
log "creating Firestore composite indexes for the page mirror"

# Pages in one section, newest first. Backs list_pages and the scoped form of
# search_pages.
gcloud firestore indexes composite create \
  --collection-group=pages \
  --field-config=field-path=sectionId,order=ascending \
  --field-config=field-path=lastModified,order=descending \
  --database='(default)' --project="$PROJECT" --async --quiet >/dev/null 2>&1 || true

# Sections to visit this sync run, least recently synced first. A
# budget-bounded run round-robins on this rather than starving the tail.
gcloud firestore indexes composite create \
  --collection-group=sections \
  --field-config=field-path=mirrored,order=ascending \
  --field-config=field-path=pagesSyncedThrough,order=ascending \
  --database='(default)' --project="$PROJECT" --async --quiet >/dev/null 2>&1 || true

# Children of one container, by name. Two collections, same shape, because
# Graph exposes sections and section groups as separate relationships and
# list_sections returns them as one tagged list.
for collection in sections sectionGroups; do
  gcloud firestore indexes composite create \
    --collection-group="$collection" \
    --field-config=field-path=parentId,order=ascending \
    --field-config=field-path=displayName,order=ascending \
    --database='(default)' --project="$PROJECT" --async --quiet >/dev/null 2>&1 || true
done

# An index exemption on the one big string. Page HTML is stored untrimmed and
# is never queried -- it is fetched by document key -- so indexing it costs
# write throughput and storage for no query at all. Firestore's index entries
# also count toward the 1 MB document limit, which is what the mirror's spill
# threshold is sized against.
gcloud firestore indexes fields update html \
  --collection-group=pageContent \
  --disable-indexes \
  --database='(default)' --project="$PROJECT" --async --quiet >/dev/null 2>&1 || true

# Confirming the exemption applied is not obvious, so the recipe is here.
# `indexes fields describe html --collection-group=pageContent` prints an
# ANCESTOR_FIELD row whether or not the exemption took, so that row says
# nothing. The discriminator is the INDEXES list: an un-exempted field carries
# three entries (ASCENDING, DESCENDING, CONTAINS) and an exempted one carries
# none. Verified 2026-08-19 against this project by describing pageContent.html
# beside pageContent.bytes, tokencache.cache and pages.title -- the three
# un-exempted fields all showed the three entries and html showed an absent
# `indexes` key. Use --format=json; the table view renders an empty list as a
# blank row, which reads like a rendering glitch.

log "indexes submitted; they build in the background. Check with:"
log "  gcloud firestore indexes composite list --database='(default)' --project=$PROJECT"
log "Every one must read READY before MIRROR_READ_ENABLED is turned on -- a query"
log "against a CREATING index fails with FAILED_PRECONDITION."

# ---------------------------------------------------------------------------
# Workload Identity Federation
#
# WIF resource names use the numeric project number, not the project ID.
# ---------------------------------------------------------------------------
PROJECT_NUMBER="$(gcloud projects describe "$PROJECT" --format='value(projectNumber)')"
[[ -n "$PROJECT_NUMBER" ]] || die "Could not resolve project number for $PROJECT"

POOL_RESOURCE="projects/${PROJECT_NUMBER}/locations/global/workloadIdentityPools/${POOL_ID}"
WIF_PROVIDER="${POOL_RESOURCE}/providers/${PROVIDER_ID}"
PRINCIPAL_SET="principalSet://iam.googleapis.com/${POOL_RESOURCE}/attribute.repository/${GITHUB_REPO}"

# A deleted pool is soft-deleted for 30 days. Re-creating one under the same
# id inside that window fails until it is undeleted or purged.
ensure_resource "workload identity pool $POOL_ID" \
  gcloud iam workload-identity-pools describe "$POOL_ID" \
    --location=global --project="$PROJECT" \
  -- \
  gcloud iam workload-identity-pools create "$POOL_ID" \
    --location=global \
    --display-name="GitHub Actions" \
    --project="$PROJECT"

# attribute.repository must be mapped, not just conditioned on: the condition
# below and the principalSet further down both resolve through it.
ATTR_MAPPING="google.subject=assertion.sub,attribute.repository=assertion.repository"
ATTR_CONDITION="assertion.repository == \"${GITHUB_REPO}\""

# This is the one resource that is updated rather than skipped when it already
# exists. A provider created earlier without the attribute condition would let
# any GitHub repository impersonate the deploy service account, and skipping
# would leave that in place silently.
if gcloud iam workload-identity-pools providers describe "$PROVIDER_ID" \
     --location=global --workload-identity-pool="$POOL_ID" \
     --project="$PROJECT" >/dev/null 2>&1; then
  log "OIDC provider $PROVIDER_ID exists, re-applying attribute mapping and condition"
  gcloud iam workload-identity-pools providers update-oidc "$PROVIDER_ID" \
    --location=global \
    --workload-identity-pool="$POOL_ID" \
    --attribute-mapping="$ATTR_MAPPING" \
    --attribute-condition="$ATTR_CONDITION" \
    --project="$PROJECT"
else
  log "creating OIDC provider $PROVIDER_ID"
  gcloud iam workload-identity-pools providers create-oidc "$PROVIDER_ID" \
    --location=global \
    --workload-identity-pool="$POOL_ID" \
    --display-name="GitHub OIDC" \
    --issuer-uri="https://token.actions.githubusercontent.com" \
    --attribute-mapping="$ATTR_MAPPING" \
    --attribute-condition="$ATTR_CONDITION" \
    --project="$PROJECT"
fi

log "granting roles/iam.workloadIdentityUser on $DEPLOY_SA to $GITHUB_REPO"
gcloud iam service-accounts add-iam-policy-binding "$DEPLOY_SA" \
  --member="$PRINCIPAL_SET" \
  --role="roles/iam.workloadIdentityUser" \
  --project="$PROJECT" --quiet >/dev/null

# ---------------------------------------------------------------------------
# Output block for issue #4.
#
# These variable names must stay byte-identical to the vars.* references in
# the deploy workflow (see the workflow sketch in project-spec.md).
# ---------------------------------------------------------------------------
echo
log "Provisioning complete. No service-account key was created; Workload"
log "Identity Federation replaces it, and none is needed anywhere."
echo
log "Run these from a clone of ${GITHUB_REPO} to complete issue #4:"
echo
cat <<GH_VARS
gh variable set GCP_PROJECT --body "$PROJECT"
gh variable set GCP_REGION  --body "$REGION"
gh variable set GAR_REGION  --body "$GAR_REGION"
gh variable set WIF_PROVIDER --body "$WIF_PROVIDER"
gh variable set DEPLOY_SA    --body "$DEPLOY_SA"
gh variable set RUNTIME_SA   --body "$RUNTIME_SA"
gh variable set MIRROR_BUCKET --body "$MIRROR_BUCKET"
GH_VARS
echo
log "These values this script cannot know. Fill in the client ID from issue #1;"
log "the openssl commands generate the rest."
echo
cat <<'GH_MANUAL'
gh variable set ONENOTE_CLIENT_ID   --body "<Azure app registration client ID from #1>"
gh variable set ONENOTE_AUTHORITY   --body "https://login.microsoftonline.com/common"
gh variable set MCP_OAUTH_CLIENT_ID --body "$(openssl rand -hex 16)"
gh secret   set MCP_OAUTH_CLIENT_SECRET --body "$(openssl rand -base64 32)"
gh secret   set MCP_TOKEN_SIGNING_KEY   --body "$(openssl rand -base64 48)"
GH_MANUAL
echo
log "Save MCP_OAUTH_CLIENT_ID and MCP_OAUTH_CLIENT_SECRET somewhere you can"
log "retrieve them. Both get typed into Claude's connector settings in issue"
log "#26, and GitHub will not show you the secret again."
