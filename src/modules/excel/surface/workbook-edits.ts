import type { WorkbookData } from "@excel/core/workbook-core";
import { checkWorkbook, refuseLongSheetName } from "@excel/core/workbook-core";
import type { ExternalLinkModel } from "@excel/core/workbook-model";
/**
 * The `Workbook` operations that take values ending up in the file, checked where they are passed. Shared by the
 * Node and browser `Workbook` surfaces.
 */
import type {
  AddExternalLinkOptions,
  ProtectWorkbookOptions,
  RegisterPersonOptions
} from "@excel/core/workbook.browser";
import {
  addExternalLink as coreAddExternalLink,
  addWorksheet as coreAddWorksheet,
  protectWorkbook,
  registerPerson as coreRegisterPerson
} from "@excel/core/workbook.browser";
import type { Worksheet } from "@excel/core/worksheet";
import { checkAutoFilter } from "@excel/surface/argument-checks";
import type { AddWorksheetOptions, WorkbookProtection } from "@excel/types";
import {
  argumentChecks,
  checkProtectionOptions,
  WORKBOOK_PROTECTION_FLAGS
} from "@excel/utils/arguments";
import { checkSheetSettings } from "@excel/utils/sheet-settings-checks";
import { isRecordArgument } from "@utils/object";

/**
 * Add a worksheet and return it. The options are checked as the writers check the same fields, so a `pageSetup`,
 * `views` or `state` value the format cannot hold is refused here rather than when the workbook is written.
 */
export function addWorksheet(
  wb: WorkbookData,
  name?: string,
  options?: AddWorksheetOptions
): Worksheet {
  const check = argumentChecks("Workbook.addWorksheet");
  refuseLongSheetName(name);
  if (check.record("options", options)) {
    checkSheetSettings("Workbook.addWorksheet", options, path => `options.${path}`);
    check.boolean("options.useSharedStrings", options.useSharedStrings);
    if (options.autoFilter !== undefined && options.autoFilter !== null) {
      checkAutoFilter("Workbook.addWorksheet", "options.autoFilter", options.autoFilter);
    }
  }
  return coreAddWorksheet(wb, name, options);
}

/** Protect the workbook's structure or windows: `protect(wb, { password: "pw", lockStructure: true })`. */
export function protect(wb: WorkbookData, options?: ProtectWorkbookOptions): Promise<void>;
/**
 * @deprecated Pass the password in the options object: `protect(wb, { password, ...options })`.
 * Removed in 1.0.
 */
export function protect(
  wb: WorkbookData,
  password: string | undefined,
  options?: Partial<WorkbookProtection>
): Promise<void>;
export function protect(
  wb: WorkbookData,
  passwordOrOptions?: string | ProtectWorkbookOptions,
  legacyOptions?: Partial<WorkbookProtection>
): Promise<void> {
  const flags = WORKBOOK_PROTECTION_FLAGS;
  try {
    checkWorkbook(wb);
    if (isRecordArgument(passwordOrOptions)) {
      checkProtectionOptions(
        "Workbook.protect",
        passwordOrOptions.password,
        passwordOrOptions,
        flags
      );
    } else {
      checkProtectionOptions("Workbook.protect", passwordOrOptions, legacyOptions, flags);
    }
  } catch (error) {
    return Promise.reject(error);
  }
  return protectWorkbook(wb, passwordOrOptions as string | undefined, legacyOptions);
}

/**
 * Register a threaded-comment author and return their person id. Registering the same display name and user id
 * again returns the existing id.
 */
export function registerPerson(
  wb: WorkbookData,
  displayName: string,
  options?: RegisterPersonOptions
): string;
/**
 * @deprecated Pass the identity in the options object: `registerPerson(wb, name, { userId, providerId })`.
 * Removed in 1.0.
 */
export function registerPerson(
  wb: WorkbookData,
  displayName: string,
  userId?: string,
  providerId?: string
): string;
export function registerPerson(
  wb: WorkbookData,
  displayName: string,
  userIdOrOptions?: string | RegisterPersonOptions,
  legacyProviderId?: string
): string {
  const check = argumentChecks("Workbook.registerPerson");
  check.required("displayName", displayName);
  check.string("displayName", displayName);
  // The deprecated positional form stores its identity as given, as it always did.
  if (isRecordArgument(userIdOrOptions)) {
    check.string("options.userId", userIdOrOptions.userId);
    check.string("options.providerId", userIdOrOptions.providerId);
  }
  return coreRegisterPerson(
    wb,
    displayName,
    userIdOrOptions as string | undefined,
    legacyProviderId
  );
}

/**
 * Declare a link to another workbook and return it. `target` is the file — a bare name or a path relative to this
 * workbook (`"data/ref.xlsx"`) — and must be a non-empty string: it becomes the relationship's target.
 *
 * The link returned is the workbook's own, as `getExternalLink` and `listExternalLinks` return it: an edit to it —
 * filling in `cachedValues`, say — reaches the file, and is checked when the workbook is written.
 */
export function addExternalLink(
  wb: WorkbookData,
  input: AddExternalLinkOptions
): ExternalLinkModel {
  const check = argumentChecks("Workbook.addExternalLink");
  check.required("input", input);
  if (check.record("input", input)) {
    check.required("input.target", input.target);
    check.string("input.target", input.target, { nonEmpty: true });
    check.oneOf("input.targetMode", input.targetMode, ["External", "Internal"]);
    if (check.array("input.sheetNames", input.sheetNames)) {
      input.sheetNames.forEach((sheet, i) =>
        check.string(`input.sheetNames[${i}]`, sheet, { nonEmpty: true })
      );
    }
    check.record("input.cachedValues", input.cachedValues);
  }
  return coreAddExternalLink(wb, input);
}
