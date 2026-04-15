import { BaseXform } from "@excel/xlsx/xform/base-xform";

interface WorkbookProtectionModel {
  lockStructure?: boolean;
  lockWindows?: boolean;
  lockRevision?: boolean;
  workbookPassword?: string;
  revisionsPassword?: string;
  algorithmName?: string;
  hashValue?: string;
  saltValue?: string;
  spinCount?: number;
}

class WorkbookProtectionXform extends BaseXform {
  get tag(): string {
    return "workbookProtection";
  }

  render(xmlStream: any, model?: WorkbookProtectionModel): void {
    if (!model) {
      return;
    }

    const attributes: Record<string, string | undefined> = {};

    if (model.lockStructure) {
      attributes.lockStructure = "1";
    }
    if (model.lockWindows) {
      attributes.lockWindows = "1";
    }
    if (model.lockRevision) {
      attributes.lockRevision = "1";
    }
    if (model.workbookPassword) {
      attributes.workbookPassword = model.workbookPassword;
    }
    if (model.revisionsPassword) {
      attributes.revisionsPassword = model.revisionsPassword;
    }
    if (model.algorithmName) {
      attributes.workbookAlgorithmName = model.algorithmName;
      attributes.workbookHashValue = model.hashValue;
      attributes.workbookSaltValue = model.saltValue;
      attributes.workbookSpinCount = model.spinCount?.toString();
    }

    if (Object.values(attributes).some(v => v !== undefined)) {
      xmlStream.leafNode(this.tag, attributes);
    }
  }

  parseOpen(node: any): boolean {
    switch (node.name) {
      case this.tag: {
        const a = node.attributes;
        this.model = {
          lockStructure: a.lockStructure === "1" || undefined,
          lockWindows: a.lockWindows === "1" || undefined,
          lockRevision: a.lockRevision === "1" || undefined,
          workbookPassword: a.workbookPassword || undefined,
          revisionsPassword: a.revisionsPassword || undefined
        } as WorkbookProtectionModel;

        if (a.workbookAlgorithmName) {
          this.model.algorithmName = a.workbookAlgorithmName;
          this.model.hashValue = a.workbookHashValue;
          this.model.saltValue = a.workbookSaltValue;
          this.model.spinCount = a.workbookSpinCount
            ? parseInt(a.workbookSpinCount, 10)
            : undefined;
        }

        return true;
      }
      default:
        return false;
    }
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }
}

export { WorkbookProtectionXform };
