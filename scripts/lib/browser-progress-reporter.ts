/**
 * Report a browser shard's progress to the process that started it, over the IPC channel.
 *
 * `scripts/run-browser-tests.ts` asks for this reporter when it watches for a stalled shard. A tester iframe can stop
 * without the browser dying — seen in WebKit on CI, a different file each time — and then nothing ends the run:
 * `testTimeout` is enforced inside the stopped iframe, and Vitest's heartbeat only notices a browser that stops
 * answering, which this one does not. So the watchdog needs to know which files the shard holds and when each
 * finishes, to name the ones that never reported — and when each *test* starts and ends, so that a stall names the
 * test that was running when the iframe stopped, not only its file. Reading that off the terminal output would tie it
 * to a reporter's wording and colours; these events are what it needs and nothing else.
 *
 * `vitest.browser.config.ts` installs it when `BROWSER_PROGRESS_IPC` is set, beside the default reporters. Without an
 * IPC channel it does nothing.
 */
import path from "node:path";
import process from "node:process";

import type { Reporter, TestCase, TestModule } from "vitest/node";

export type BrowserProgressMessage =
  | { readonly type: "queued"; readonly file: string }
  | { readonly type: "end"; readonly file: string }
  | { readonly type: "test-start"; readonly file: string; readonly test: string }
  | { readonly type: "test-end"; readonly file: string; readonly test: string };

function relativeFile(moduleId: string): string {
  return path.relative(process.cwd(), moduleId).split(path.sep).join("/");
}

function send(message: BrowserProgressMessage): void {
  process.send?.(message);
}

export class BrowserProgressReporter implements Reporter {
  onInit(): void {
    // An open IPC channel keeps a process alive; Vitest must still exit when its run ends.
    process.channel?.unref();
  }

  // Queued, not `onTestRunStart`: in Vitest 5.1.0 the run-start list is the whole suite, before `--shard` applies.
  onTestModuleQueued(testModule: TestModule): void {
    send({ type: "queued", file: relativeFile(testModule.moduleId) });
  }

  onTestModuleEnd(testModule: TestModule): void {
    send({ type: "end", file: relativeFile(testModule.moduleId) });
  }

  onTestCaseReady(testCase: TestCase): void {
    send({
      type: "test-start",
      file: relativeFile(testCase.module.moduleId),
      test: testCase.fullName
    });
  }

  onTestCaseResult(testCase: TestCase): void {
    send({
      type: "test-end",
      file: relativeFile(testCase.module.moduleId),
      test: testCase.fullName
    });
  }
}
