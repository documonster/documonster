/**
 * Browser half of the early-exit suite: the browser build of the ZIP parser over a Web
 * `ReadableStream` and an async generator. See `unzip/early-exit.shared.ts`.
 */

import {
  PORTABLE_SOURCES,
  defineTarEarlyExitSuite,
  defineZipEarlyExitSuite
} from "@archive/__tests__/unzip/early-exit.shared";

defineZipEarlyExitSuite(PORTABLE_SOURCES);
defineTarEarlyExitSuite(PORTABLE_SOURCES);
