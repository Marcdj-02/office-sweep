# pptx-sweep

A utility to remove artifacts from PowerPoint files, such as comments, notes, and metadata.

[![npm version](https://badge.fury.io/js/pptx-sweep.svg)](https://badge.fury.io/js/pptx-sweep)

## How to use

```typescript
import { pptxSweep } from "pptx-sweep";

const sourceFile = "powerpoint.pptx";
const targetFile = "powerpoint-swept.pptx";

await pptxSweep(sourceFile, targetFile, {
  remove: {
    totalTime: true,
    core: {
      title: true,
      creator: true,
      lastMofiiedBy: true,
      revision: true,
      created: true,
      modified: true,
    },
    thumbnail: true,
    notes: true,
    comments: {
      modern: true,
      legacy: true,
    },
    authors: true,
    view: true,
  },
});
```

## Options

```typescript
type SweepOptions = {
  remove: {
    totalTime?: boolean;
    core: {
      title?: boolean;
      creator?: boolean;
      lastMofiiedBy?: boolean;
      revision?: boolean;
      created?: boolean;
      modified?: boolean;
    };
    thumbnail?: boolean;
    notes?: boolean;
    comments?: {
      modern?: boolean;
      legacy?: boolean;
    };
    authors?: boolean;
    view?: boolean;
  };
};
```

| Option                      | Description                                                       |
| --------------------------- | ----------------------------------------------------------------- |
| `remove.totalTime`          | Remove the total time of the presentation.                        |
| `remove.core.title`         | Remove the title of the presentation.                             |
| `remove.core.creator`       | Remove the name of the person who created the presentation.       |
| `remove.core.lastMofiiedBy` | Remove the name of the person who last modified the presentation. |
| `remove.core.revision`      | Remove the revision number.                                       |
| `remove.core.created`       | Remove the date and time the presentation was created.            |
| `remove.core.modified`      | Remove the date and time the presentation was last modified.      |
| `remove.thumbnail`          | Remove the thumbnail image.                                       |
| `remove.notes`              | Remove the (presenter) notes.                                     |
| `remove.comments.modern`    | Remove the modern (2018 spec) comments.                           |
| `remove.comments.legacy`    | Remove the legacy (2006spec) comments.                            |
| `remove.authors`            | Remove the list of authors.                                       |
| `remove.view`               | Remove the view settings.                                         |
