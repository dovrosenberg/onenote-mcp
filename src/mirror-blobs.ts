// Cloud Storage for the two things that do not fit in a Firestore document.
//
// Rendered ink PNGs, the InkML they were rendered from, and any page HTML over
// HTML_INLINE_LIMIT_BYTES. Firestore caps a document at 1 MiB and MAX_INK_PNG_BYTES is
// already 750 KB, so a PNG would fill most of a document on its own and leave no room to
// grow if that budget is ever raised.
//
// **The InkML is stored beside the PNG, and that is not redundant.** docs/content.md records
// MAX_INK_PNG_BYTES as "a budget chosen rather than measured against any client's cap",
// and `fitInkToByteBudget` shrinks a render by re-rasterising the SVG and measuring. If
// only the PNG were kept, changing that number would mean re-fetching every inked page
// from Graph — hours of the request budget to correct a guess. Keeping the InkML makes a
// re-render a local operation. It is the same protection the raw HTML gets by being
// stored untrimmed.
//
// Nothing in this file has an automated test, and it cannot get one here: what is at
// stake is whether the runtime service account can actually write the bucket and whether
// an object survives a delete-and-rewrite, neither of which a fake observes. Everything
// that could be *decided* wrongly — the object names, and which kinds exist — lives in
// ./mirror-schema.ts and is tested there. Verify this file by running a sync and looking
// in the bucket.

import { Storage, type Bucket } from '@google-cloud/storage';

import { htmlObjectName, inkObjectName, inkmlObjectName } from './mirror-schema.ts';

/** Content types, set on write so a browser opening an object from the console renders it. */
const PNG = 'image/png';
const INKML = 'application/inkml+xml';
const HTML = 'text/html; charset=utf-8';

/**
 * The slice of the blob store the sync writes through. Declared narrowly so
 * ./mirror-sync.ts can be tested against a plain object that records calls.
 */
export interface MirrorBlobWriter {
  putInk(pageId: string, png: Uint8Array): Promise<string>;
  putInkml(pageId: string, inkml: string): Promise<string>;
  putHtml(pageId: string, html: string): Promise<string>;
  /** Every object belonging to one page. Missing objects are not an error. */
  deleteForPage(pageId: string): Promise<void>;
}

/** The slice the read path uses. */
export interface MirrorBlobReader {
  /** Null when the object is absent, which a caller treats as a mirror miss. */
  getInk(pageId: string): Promise<Uint8Array | null>;
  getHtml(pageId: string): Promise<string | null>;
}

export interface MirrorBlobStore extends MirrorBlobWriter, MirrorBlobReader {
  getInkml(pageId: string): Promise<string | null>;
}

/**
 * A GCS error's status code, if it carries one.
 *
 * The library throws `ApiError` with a numeric `code`, but a network failure below it
 * throws something else entirely, so this reads defensively rather than casting.
 */
function statusOf(err: unknown): number | undefined {
  const code = (err as { code?: unknown }).code;
  return typeof code === 'number' ? code : undefined;
}

class GcsMirrorBlobStore implements MirrorBlobStore {
  readonly #bucket: Bucket;

  constructor(bucket: Bucket) {
    this.#bucket = bucket;
  }

  async putInk(pageId: string, png: Uint8Array): Promise<string> {
    return this.#put(inkObjectName(pageId), Buffer.from(png), PNG);
  }

  async putInkml(pageId: string, inkml: string): Promise<string> {
    return this.#put(inkmlObjectName(pageId), Buffer.from(inkml, 'utf8'), INKML);
  }

  async putHtml(pageId: string, html: string): Promise<string> {
    return this.#put(htmlObjectName(pageId), Buffer.from(html, 'utf8'), HTML);
  }

  async getInk(pageId: string): Promise<Uint8Array | null> {
    const bytes = await this.#get(inkObjectName(pageId));
    return bytes === null ? null : new Uint8Array(bytes);
  }

  async getInkml(pageId: string): Promise<string | null> {
    const bytes = await this.#get(inkmlObjectName(pageId));
    return bytes === null ? null : bytes.toString('utf8');
  }

  async getHtml(pageId: string): Promise<string | null> {
    const bytes = await this.#get(htmlObjectName(pageId));
    return bytes === null ? null : bytes.toString('utf8');
  }

  /**
   * Every object for one page, built from the id rather than read from the document.
   *
   * A 404 on any of them is success: a typed page has no ink, a small page has no
   * spilled HTML, and a partially-written page from a failed run has some subset. The
   * caller is deleting the page, so "it was not there" and "it is not there now" are the
   * same outcome.
   */
  async deleteForPage(pageId: string): Promise<void> {
    await Promise.all(
      [inkObjectName(pageId), inkmlObjectName(pageId), htmlObjectName(pageId)].map(
        async (name) => {
          try {
            await this.#bucket.file(name).delete();
          } catch (err) {
            if (statusOf(err) === 404) return;
            throw err;
          }
        },
      ),
    );
  }

  async #put(name: string, body: Buffer, contentType: string): Promise<string> {
    // resumable: false because every object here is well under the 5 MB threshold where
    // a resumable upload starts paying for itself, and a resumable session costs an
    // extra round trip per object — which, across a 2000-page backfill, is 2000 round
    // trips for nothing.
    await this.#bucket.file(name).save(body, {
      resumable: false,
      contentType,
    });
    return name;
  }

  async #get(name: string): Promise<Buffer | null> {
    try {
      const [contents] = await this.#bucket.file(name).download();
      return contents;
    } catch (err) {
      // Absent is a mirror miss, which the read path answers by going to Graph. Anything
      // else propagates: a permission failure that read as "not mirrored" would send
      // every request to Graph and exhaust the hourly budget with nothing saying why.
      if (statusOf(err) === 404) return null;
      throw err;
    }
  }
}

/**
 * Build the blob store for a bucket.
 *
 * The `Storage` client is constructed per call rather than memoised the way
 * ./firestore.ts memoises Firestore, because there is exactly one call site — the wiring
 * in ./server.ts — and a second one would be a mistake worth seeing in a diff.
 * Construction opens no connection.
 */
export function createMirrorBlobStore(bucketName: string, projectId?: string): MirrorBlobStore {
  const storage = new Storage(projectId === undefined ? {} : { projectId });
  return new GcsMirrorBlobStore(storage.bucket(bucketName));
}
