// scripts/utils/batch.js
// Minimal utility for resilient Word.js batching & retries
// Uses exponential backoff for transient sync errors (e.g., coauthoring conflicts)
const DEFAULT_BACKOFFS_MS = [100, 250, 500];

export async function runBatch(batchFn, backoffs = DEFAULT_BACKOFFS_MS) {
  return Word.run(async (context) => {
    for (let attempt = 0; attempt <= backoffs.length; attempt++) {
      try {
        await batchFn(context);
        await context.sync(); // single sync per batch — recommended by MS docs
        return;
      } catch (err) {
        const isLast = attempt === backoffs.length;
        if (isLast) throw err;
        await delay(backoffs[attempt]);
      }
    }
  });
}

export function delay(ms) {
  return new Promise((res) => setTimeout(res, ms));
}
