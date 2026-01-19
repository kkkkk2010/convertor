export type DocJson = {
  schemaVersion: 1;
  slideSize: {
    width: number;
    height: number;
    unit: "in";
  };
  slides: Slide[];
};

export type Slide = {
  id: string;
  background: {
    type: "image";
    src: string;
  };
  elements: SlideElement[];
};

export type SlideElement = TextElement | ImageElement;

export type TextElement = {
  id: string;
  type: "text";
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  text: string;
  style: {
    fontFamily: string;
    fontSizePt: number;
    color: string;
    bold: boolean;
    italic: boolean;
    underline: boolean;
    align: "left" | "center" | "right" | "justify";
    lineHeight: number | null;
  };
};

export type ImageElement = {
  id: string;
  type: "image";
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  src: string;
  objectFit: "cover";
};
