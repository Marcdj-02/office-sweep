import JSZip from "jszip";
import { SweepOptions } from "../office";
import { ModifyReturn } from "../types";

const removalPatterns: Record<
  keyof NonNullable<NonNullable<NonNullable<SweepOptions["remove"]>["ppt"]>["core"]>,
  | {
      action: "remove";
      pattern: string;
    }
  | {
      action: "replace";
      pattern: string;
      replacement: string;
    }
> = {
  title: {
    action: "remove",
    pattern: "<dc:title[^>]*[^>]*>[^<]*</dc:title>",
  },
  creator: {
    action: "remove",
    pattern: "<dc:creator[^>]*[^>]*>[^<]*</dc:creator>",
  },
  lastModifiedBy: {
    action: "remove",
    pattern: "<cp:lastModifiedBy[^>]*[^>]*>[^<]*</cp:lastModifiedBy>",
  },
  revision: {
    action: "remove",
    pattern: "<cp:revision[^>]*[^>]*>[^<]*</cp:revision>",
  },
  created: {
    action: "remove",
    pattern: "<dcterms:created[^>]*[^>]*>[^<]*</dcterms:created>",
  },
  modified: {
    action: "remove",
    pattern: "<dcterms:modified[^>]*[^>]*>[^<]*</dcterms:modified>",
  },
};

export async function modifyCoreProperties(
  zip: JSZip,
  corePath: string,
  options: SweepOptions
): Promise<void> {
  let core = await zip.file(corePath)?.async("string");
  if (!core) {
    throw new Error(`File not found: ${corePath}`);
  }

  if (options.remove?.ppt?.core) {
    for (const key of Object.keys(options.remove.ppt.core)) {
      const castedKey = key as keyof typeof options.remove.ppt.core;

      if (options.remove.ppt.core[castedKey]) {
        const config = removalPatterns[castedKey];
        if (!config) {
          throw new Error(`Unknown key: ${castedKey}`);
        }

        const { action, pattern } = config;

        switch (action) {
          case "remove":
            core = core.replace(new RegExp(pattern, "g"), "");
            break;
          case "replace":
            const { replacement } = config;
            core = core.replace(new RegExp(pattern, "g"), replacement);
            break;
          default:
            throw new Error(`Unknown action: ${action}`);
        }
      }
    }
  }

  zip.file(corePath, core);
}
