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
gcloud services enable \
  run.googleapis.com \
  artifactregistry.googleapis.com \
  firestore.googleapis.com \
  iamcredentials.googleapis.com \
  sts.googleapis.com \
  cloudresourcemanager.googleapis.com \
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
