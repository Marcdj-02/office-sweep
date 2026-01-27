export type Image = {
  slideIndexes: number[];
  name: string;
  internalPath: string;
};

export type ModifyReturn = {
  images: Image[];
};
