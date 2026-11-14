import { describeAsyncDispose } from "@archive/__tests__/io/async-dispose.shared";
import { Compression } from "@archive/index";
import { describe, expect, it, vi } from "vitest";

describeAsyncDispose();

// Only a browser has workers: `using` terminates the pool there.
describe("using a compression worker pool", () => {
  it("terminates it on leaving the block", () => {
    const terminate = vi.spyOn(Compression.Workers.Pool.prototype, "terminate");
    {
      using pool = new Compression.Workers.Pool({ maxWorkers: 1 });
      expect(pool.isTerminated()).toBe(false);
    }
    expect(terminate).toHaveBeenCalledTimes(1);
    terminate.mockRestore();
  });
});
