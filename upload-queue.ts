/**
 * Concurrency control queue for upload tasks
 */
export class UploadQueue {
  private queue: Array<() => Promise<void>> = [];
  private maxConcurrency: number;
  private successCount = 0;
  private failCount = 0;
  private totalFiles: number;
  private running = 0;

  constructor(maxConcurrency: number, totalFiles: number) {
    this.maxConcurrency = maxConcurrency;
    this.totalFiles = totalFiles;
  }

  add(task: () => Promise<void>) {
    this.queue.push(task);
    this.process();
  }

  private async process() {
    while (this.running < this.maxConcurrency && this.queue.length > 0) {
      const task = this.queue.shift()!;
      this.running++;

      // Execute task and handle completion
      (async () => {
        try {
          await task();
          this.successCount++;
        } catch (error) {
          this.failCount++;
        } finally {
          this.running--;
          // Process next task when one completes
          this.process();
        }
      })();
    }
  }

  async waitForCompletion() {
    // Wait for all tasks to complete
    while (this.queue.length > 0 || this.running > 0) {
      await new Promise((resolve) => setTimeout(resolve, 50));
    }
  }

  getStats() {
    return {
      success: this.successCount,
      failed: this.failCount,
      total: this.totalFiles,
    };
  }
}
