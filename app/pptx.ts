export type TemplateId = "simple" | "cool" | "fashion";
export type LogoId =
  | "rookies"
  | "gp-2024"
  | "gp-2025"
  | "gp-2026"
  | "guild"
  | "world-series"
  | null;

export type FacePptData = {
  name: string;
  title: string;
  count: 3 | 4;
  template: TemplateId;
  blocks: Array<{
    title: string;
    description: string;
    photo: {
      dataUrl: string;
      name: string;
      zoom: number;
      x: number;
      y: number;
    } | null;
    rookiesLogoId: LogoId;
  }>;
};

export type MotivationPptData = {
  name: string;
  title: string;
  count: 5 | 6 | 7;
  template: TemplateId;
  motivationUpSummary: string;
  motivationDownSummary: string;
  episodes: Array<{
    period: string;
    title: string;
    description: string;
    motivation: number;
    rookiesLogoId: LogoId;
  }>;
};

type Palette = {
  bg: string;
  fg: string;
  accent: string;
  soft: string;
  muted: string;
};

const palettes: Record<TemplateId, Palette> = {
  simple: {
    bg: "FFFEF9",
    fg: "20201F",
    accent: "D8FF4F",
    soft: "F0EFE8",
    muted: "74756F",
  },
  cool: {
    bg: "1D2538",
    fg: "F8F8F4",
    accent: "8FA7FF",
    soft: "2B354D",
    muted: "AAB4C6",
  },
  fashion: {
    bg: "F5EEE6",
    fg: "243F37",
    accent: "FF7548",
    soft: "E4D8CA",
    muted: "6C746F",
  },
};

const logoNames: Record<Exclude<LogoId, null>, string> = {
  rookies: "ROOKIES",
  "gp-2024": "CAREER ROOKIES GP 2024",
  "gp-2025": "CAREER ROOKIES GP 2025",
  "gp-2026": "CAREER ROOKIES GP 2026",
  guild: "ROOKIES GUILD",
  "world-series": "ROOKIES WORLD SERIES",
};

const logoAssetPaths: Partial<Record<Exclude<LogoId, null>, string>> = {
  rookies: "/logos/rookies/rookies.png",
  "gp-2024": "/logos/rookies/gp-2024.png",
  "gp-2025": "/logos/rookies/gp-2025.png",
  "gp-2026": "/logos/rookies/gp-2026.png",
  guild: "/logos/rookies/guild.png",
  "world-series": "/logos/rookies/world-series.png",
};

const sanitizeFileName = (value: string) =>
  value.replace(/[\\/:*?"<>|]/g, "_").trim() || "MY STORY MAKER";

type LogoAsset = {
  data: string;
  ratio: number;
};

const placeholderLogo = (
  id: Exclude<LogoId, null>,
  palette: Palette,
): LogoAsset => {
  const canvas = document.createElement("canvas");
  canvas.width = 720;
  canvas.height = 150;
  const context = canvas.getContext("2d");
  if (!context) return { data: "", ratio: 4.8 };
  context.clearRect(0, 0, canvas.width, canvas.height);
  context.strokeStyle = `#${palette.fg}`;
  context.lineWidth = 6;
  context.strokeRect(7, 7, 706, 136);
  context.fillStyle = `#${palette.fg}`;
  context.font = "900 52px Arial";
  context.textAlign = "center";
  context.textBaseline = "middle";
  context.fillText(logoNames[id], 360, 78);
  return { data: canvas.toDataURL("image/png"), ratio: 4.8 };
};

const logoDataUri = async (
  id: Exclude<LogoId, null>,
  palette: Palette,
): Promise<LogoAsset> => {
  const path = logoAssetPaths[id];
  if (!path) return placeholderLogo(id, palette);
  try {
    const image = new Image();
    image.src = path;
    await image.decode();
    const canvas = document.createElement("canvas");
    canvas.width = image.naturalWidth;
    canvas.height = image.naturalHeight;
    const context = canvas.getContext("2d");
    if (!context) return placeholderLogo(id, palette);
    context.drawImage(image, 0, 0);
    if (palette.bg === palettes.cool.bg && id !== "gp-2024") {
      context.globalCompositeOperation = "source-in";
      context.fillStyle = `#${palette.fg}`;
      context.fillRect(0, 0, canvas.width, canvas.height);
    }
    return {
      data: canvas.toDataURL("image/png"),
      ratio: image.naturalWidth / image.naturalHeight,
    };
  } catch {
    return placeholderLogo(id, palette);
  }
};

const fitLogo = (
  asset: LogoAsset,
  x: number,
  y: number,
  maxWidth: number,
  maxHeight: number,
) => {
  const width = Math.min(maxWidth, maxHeight * asset.ratio);
  const height = width / asset.ratio;
  return {
    x,
    y: y + (maxHeight - height) / 2,
    w: width,
    h: height,
  };
};

const cropImageData = async (
  source: string,
  width: number,
  height: number,
  zoom: number,
  x: number,
  y: number,
) => {
  const image = new Image();
  image.src = source;
  await image.decode();
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(width * 160));
  canvas.height = Math.max(1, Math.round(height * 160));
  const context = canvas.getContext("2d");
  if (!context) return source;
  const baseScale = Math.max(
    canvas.width / image.width,
    canvas.height / image.height,
  );
  const scale = baseScale * zoom;
  const drawWidth = image.width * scale;
  const drawHeight = image.height * scale;
  const maxX = Math.max(0, drawWidth - canvas.width);
  const maxY = Math.max(0, drawHeight - canvas.height);
  const drawX = -(x / 100) * maxX;
  const drawY = -(y / 100) * maxY;
  context.drawImage(image, drawX, drawY, drawWidth, drawHeight);
  return canvas.toDataURL("image/jpeg", 0.9);
};

async function createBase(template: TemplateId) {
  const imported = await import("pptxgenjs");
  const PptxGenJS = imported.default;
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "A4_LANDSCAPE", width: 11.69, height: 8.27 });
  pptx.layout = "A4_LANDSCAPE";
  pptx.author = "MY STORY MAKER";
  pptx.company = "MY STORY MAKER";
  pptx.subject = "自己紹介資料";
  pptx.theme = {
    headFontFace: "Yu Gothic",
    bodyFontFace: "Yu Gothic",
  };
  const palette = palettes[template];
  const slide = pptx.addSlide();
  slide.background = { color: palette.bg };
  return { pptx, slide, palette };
}

function addHeader(
  slide: ReturnType<Awaited<ReturnType<typeof createBase>>["pptx"]["addSlide"]>,
  palette: Palette,
  title: string,
  name: string,
  eyebrow: string,
) {
  slide.addText(eyebrow, {
    x: 0.61,
    y: 0.42,
    w: 5.8,
    h: 0.2,
    fontFace: "Arial",
    fontSize: 7,
    bold: true,
    charSpacing: 2,
    color: palette.muted,
    margin: 0,
  });
  slide.addText(title, {
    x: 0.61,
    y: 0.69,
    w: 7.6,
    h: 0.5,
    fontFace: "Yu Gothic",
    fontSize: 23,
    bold: true,
    color: palette.fg,
    margin: 0,
    breakLine: false,
    fit: "shrink",
  });
  slide.addShape("rect", {
    x: 9.42,
    y: 0.54,
    w: 0.06,
    h: 0.42,
    line: { color: palette.accent, transparency: 100 },
    fill: { color: palette.accent },
  });
  slide.addText(name, {
    x: 9.62,
    y: 0.58,
    w: 1.45,
    h: 0.3,
    align: "right",
    fontFace: "Yu Gothic",
    fontSize: 10,
    bold: true,
    color: palette.fg,
    margin: 0,
    fit: "shrink",
  });
}

function addFooter(
  slide: ReturnType<Awaited<ReturnType<typeof createBase>>["pptx"]["addSlide"]>,
  palette: Palette,
  left: string,
) {
  slide.addText(left, {
    x: 0.61,
    y: 7.84,
    w: 1.6,
    h: 0.13,
    fontFace: "Arial",
    fontSize: 5,
    bold: true,
    charSpacing: 1.2,
    color: palette.muted,
    margin: 0,
  });
  slide.addShape("line", {
    x: 2.1,
    y: 7.9,
    w: 7.55,
    h: 0,
    line: { color: palette.muted, transparency: 68, width: 0.5 },
  });
  slide.addText("MY STORY MAKER", {
    x: 9.8,
    y: 7.84,
    w: 1.25,
    h: 0.13,
    align: "right",
    fontFace: "Arial",
    fontSize: 5,
    bold: true,
    charSpacing: 1.2,
    color: palette.muted,
    margin: 0,
  });
}

async function addDecorationSlide(
  pptx: Awaited<ReturnType<typeof createBase>>["pptx"],
  palette: Palette,
  template: TemplateId,
) {
  const slide = pptx.addSlide();
  slide.background = { color: palette.bg };

  slide.addText("DECORATION TOOLKIT", {
    x: 0.61,
    y: 0.5,
    w: 7.6,
    h: 0.46,
    fontFace: "Arial",
    fontSize: 23,
    bold: true,
    charSpacing: 1.4,
    color: palette.fg,
    margin: 0,
  });
  slide.addText(
    "好きな素材をコピーして、1枚目に貼り付けて使えます。テキスト・色・サイズも自由に編集できます。",
    {
      x: 0.61,
      y: 1.05,
      w: 8.8,
      h: 0.28,
      fontFace: "Yu Gothic",
      fontSize: 8.5,
      color: palette.muted,
      margin: 0,
      fit: "shrink",
    },
  );
  slide.addShape("roundRect", {
    x: 9.63,
    y: 0.54,
    w: 1.43,
    h: 0.4,
    rectRadius: 0.06,
    line: { color: palette.accent, transparency: 100 },
    fill: { color: palette.accent },
  });
  slide.addText(template.toUpperCase(), {
    x: 9.63,
    y: 0.66,
    w: 1.43,
    h: 0.13,
    align: "center",
    fontFace: "Arial",
    fontSize: 6.5,
    bold: true,
    charSpacing: 1.2,
    color: "20201F",
    margin: 0,
  });
  slide.addShape("line", {
    x: 0.61,
    y: 1.52,
    w: 10.47,
    h: 0,
    line: { color: palette.muted, transparency: 72, width: 0.6 },
  });

  const addSectionLabel = (text: string, x: number, width: number) => {
    slide.addText(text, {
      x,
      y: 1.79,
      w: width,
      h: 0.18,
      fontFace: "Arial",
      fontSize: 7,
      bold: true,
      charSpacing: 1.5,
      color: palette.muted,
      margin: 0,
    });
  };

  addSectionLabel("TEXT LABELS", 0.61, 3.15);
  addSectionLabel("NUMBERS / SHAPES / LINES", 4.07, 3.25);
  addSectionLabel("SPEECH BUBBLES", 7.73, 3.35);

  const textLabels = [
    { text: "MY STORY", fill: palette.accent, color: "20201F" },
    { text: "MY CHALLENGE", fill: palette.soft, color: palette.fg },
    { text: "FAVORITE", fill: palette.fg, color: palette.bg },
  ];
  textLabels.forEach((item, index) => {
    const y = 2.19 + index * 0.55;
    slide.addShape("roundRect", {
      x: 0.61,
      y,
      w: 2.75,
      h: 0.38,
      rectRadius: 0.06,
      line: { color: item.fill, transparency: 100 },
      fill: { color: item.fill },
    });
    slide.addText(item.text, {
      x: 0.8,
      y: y + 0.12,
      w: 2.37,
      h: 0.12,
      fontFace: "Arial",
      fontSize: 7,
      bold: true,
      charSpacing: 1.2,
      color: item.color,
      margin: 0,
    });
  });
  slide.addText("Highlight!", {
    x: 0.61,
    y: 4.04,
    w: 2.95,
    h: 0.52,
    fontFace: "Arial",
    fontSize: 27,
    bold: true,
    italic: true,
    color: palette.fg,
    margin: 0,
    fit: "shrink",
  });
  slide.addShape("line", {
    x: 0.63,
    y: 4.65,
    w: 2.65,
    h: -0.12,
    line: { color: palette.accent, width: 6, transparency: 12 },
  });
  slide.addText("YOUR WORDS HERE", {
    x: 0.61,
    y: 5.08,
    w: 2.9,
    h: 0.38,
    fontFace: "Yu Gothic",
    fontSize: 13,
    bold: true,
    color: palette.fg,
    margin: 0,
    fit: "shrink",
  });
  slide.addText("見出しや短いメッセージに", {
    x: 0.61,
    y: 5.51,
    w: 2.9,
    h: 0.2,
    fontFace: "Yu Gothic",
    fontSize: 7,
    color: palette.muted,
    margin: 0,
  });

  [1, 2, 3, 4].forEach((value, index) => {
    const x = 4.07 + index * 0.74;
    slide.addShape("ellipse", {
      x,
      y: 2.17,
      w: 0.54,
      h: 0.54,
      line: { color: palette.accent, transparency: 100 },
      fill: { color: index % 2 === 0 ? palette.accent : palette.soft },
    });
    slide.addText(String(value).padStart(2, "0"), {
      x,
      y: 2.36,
      w: 0.54,
      h: 0.12,
      align: "center",
      fontFace: "Arial",
      fontSize: 7,
      bold: true,
      color: index % 2 === 0 ? "20201F" : palette.fg,
      margin: 0,
    });
  });

  const shapeTypes = ["ellipse", "rect", "roundRect", "diamond", "triangle"] as const;
  shapeTypes.forEach((shapeType, index) => {
    const x = 4.07 + index * 0.64;
    slide.addShape(shapeType, {
      x,
      y: 3.17,
      w: 0.42,
      h: 0.42,
      line: {
        color: index === 4 ? palette.accent : palette.fg,
        width: 1.1,
      },
      fill: {
        color: index === 4 ? palette.accent : palette.soft,
        transparency: index === 4 ? 0 : 18,
      },
    });
  });
  slide.addShape("line", {
    x: 4.07,
    y: 4.13,
    w: 2.95,
    h: 0,
    line: {
      color: palette.fg,
      width: 1.5,
      beginArrowType: "none",
      endArrowType: "triangle",
    },
  });
  slide.addShape("line", {
    x: 4.07,
    y: 4.65,
    w: 2.95,
    h: 0,
    line: {
      color: palette.muted,
      width: 1.2,
      dashType: "dash",
    },
  });
  slide.addShape("chevron", {
    x: 4.07,
    y: 5.12,
    w: 0.72,
    h: 0.48,
    line: { color: palette.accent, transparency: 100 },
    fill: { color: palette.accent },
  });
  slide.addShape("star5", {
    x: 4.96,
    y: 5.06,
    w: 0.58,
    h: 0.58,
    line: { color: palette.fg, transparency: 100 },
    fill: { color: palette.fg },
  });
  slide.addText("↗  +  #", {
    x: 5.8,
    y: 5.15,
    w: 1.2,
    h: 0.3,
    fontFace: "Arial",
    fontSize: 18,
    bold: true,
    color: palette.accent,
    margin: 0,
  });

  const speechBubbles = [
    {
      shape: "wedgeRoundRectCallout" as const,
      text: "ひとこと",
      fill: palette.soft,
      color: palette.fg,
      flipH: false,
    },
    {
      shape: "wedgeEllipseCallout" as const,
      text: "POINT!",
      fill: palette.accent,
      color: "20201F",
      flipH: true,
    },
    {
      shape: "wedgeRectCallout" as const,
      text: "コメント",
      fill: palette.fg,
      color: palette.bg,
      flipH: true,
    },
    {
      shape: "cloudCallout" as const,
      text: "アイデア",
      fill: palette.soft,
      color: palette.fg,
      flipH: false,
    },
    {
      shape: "accentBorderCallout1" as const,
      text: "補足を入力",
      fill: palette.bg,
      color: palette.fg,
      flipH: false,
    },
    {
      shape: "borderCallout2" as const,
      text: "自由に編集",
      fill: palette.bg,
      color: palette.fg,
      flipH: true,
    },
  ];
  speechBubbles.forEach((bubble, index) => {
    const column = index % 2;
    const row = Math.floor(index / 2);
    const x = 7.73 + column * 1.73;
    const y = 2.12 + row * 1.27;
    slide.addText(bubble.text, {
      shape: bubble.shape,
      x,
      y,
      w: 1.58,
      h: 0.94,
      flipH: bubble.flipH,
      align: "center",
      valign: "middle",
      fontFace: bubble.text === "POINT!" ? "Arial" : "Yu Gothic",
      fontSize: bubble.text === "POINT!" ? 9 : 7.5,
      bold: true,
      color: bubble.color,
      margin: [4, 6, 10, 6],
      line: {
        color: index < 4 ? palette.fg : palette.accent,
        transparency: index < 4 ? 72 : 0,
        width: index < 4 ? 0.8 : 1.4,
      },
      fill: {
        color: bubble.fill,
        transparency: index >= 4 ? 100 : 0,
      },
      fit: "shrink",
    });
  });

  slide.addShape("line", {
    x: 0.61,
    y: 6.27,
    w: 10.47,
    h: 0,
    line: { color: palette.muted, transparency: 72, width: 0.6 },
  });
  slide.addText("COLOR PALETTE", {
    x: 0.61,
    y: 6.56,
    w: 1.7,
    h: 0.18,
    fontFace: "Arial",
    fontSize: 7,
    bold: true,
    charSpacing: 1.5,
    color: palette.muted,
    margin: 0,
  });
  [
    palette.accent,
    palette.fg,
    palette.soft,
    palette.muted,
    palette.bg,
  ].forEach((color, index) => {
    slide.addShape("ellipse", {
      x: 2.38 + index * 0.55,
      y: 6.43,
      w: 0.34,
      h: 0.34,
      line: { color: palette.muted, transparency: 58, width: 0.6 },
      fill: { color },
    });
  });
  slide.addText("コピー＆ペーストで、あなたらしい一枚に。", {
    x: 6.26,
    y: 6.51,
    w: 4.82,
    h: 0.22,
    align: "right",
    fontFace: "Yu Gothic",
    fontSize: 8,
    bold: true,
    color: palette.fg,
    margin: 0,
  });

  addFooter(slide, palette, "EDITABLE ASSETS");
}

export async function createFacePowerPoint(data: FacePptData) {
  const { pptx, slide, palette } = await createBase(data.template);
  pptx.title = data.title;
  addHeader(
    slide,
    palette,
    data.title,
    data.name,
    `ABOUT ME / ${String(data.count).padStart(2, "0")} FACES`,
  );

  const count = data.count;
  const totalWidth = 10.47;
  const gap = count === 4 ? 0.16 : 0.24;
  const cardWidth = (totalWidth - gap * (count - 1)) / count;
  const xStart = 0.61;
  const photoY = 1.45;
  const photoHeight = 3.58;
  const textY = 5.03;
  const textHeight = 2.55;

  for (const [index, block] of data.blocks.slice(0, count).entries()) {
    const x = xStart + index * (cardWidth + gap);
    slide.addShape("rect", {
      x,
      y: photoY,
      w: cardWidth,
      h: photoHeight + textHeight,
      line: { color: palette.soft, transparency: 100 },
      fill: { color: palette.soft },
    });
    if (block.photo) {
      const cropped = await cropImageData(
        block.photo.dataUrl,
        cardWidth,
        photoHeight,
        block.photo.zoom,
        block.photo.x,
        block.photo.y,
      );
      slide.addImage({
        data: cropped,
        x,
        y: photoY,
        w: cardWidth,
        h: photoHeight,
      });
    } else {
      slide.addShape("rect", {
        x,
        y: photoY,
        w: cardWidth,
        h: photoHeight,
        line: { color: palette.muted, transparency: 84 },
        fill: { color: palette.muted, transparency: 38 },
      });
      slide.addText("PHOTO", {
        x,
        y: photoY + 1.58,
        w: cardWidth,
        h: 0.26,
        align: "center",
        fontFace: "Arial",
        fontSize: 9,
        bold: true,
        charSpacing: 2,
        color: palette.bg,
        transparency: 20,
        margin: 0,
      });
    }
    slide.addText(`0${index + 1}`, {
      x: x + cardWidth - 0.68,
      y: photoY + photoHeight - 0.55,
      w: 0.58,
      h: 0.48,
      align: "right",
      fontFace: "Arial",
      fontSize: 25,
      bold: true,
      color: palette.accent,
      margin: 0,
    });

    let contentTop = textY + 0.18;
    if (block.rookiesLogoId) {
      const logo = await logoDataUri(block.rookiesLogoId, palette);
      slide.addImage({
        data: logo.data,
        ...fitLogo(
          logo,
          x + 0.18,
          contentTop,
          Math.min(cardWidth - 0.36, 1.48),
          0.3,
        ),
      });
      contentTop += 0.39;
    }
    slide.addText(block.title, {
      x: x + 0.18,
      y: contentTop,
      w: cardWidth - 0.36,
      h: 0.37,
      fontFace: "Yu Gothic",
      fontSize: count === 4 ? 11 : 13,
      bold: true,
      color: palette.fg,
      margin: 0,
      fit: "shrink",
    });
    slide.addText(block.description, {
      x: x + 0.18,
      y: contentTop + 0.49,
      w: cardWidth - 0.36,
      h: textY + textHeight - contentTop - 0.72,
      fontFace: "Yu Gothic",
      fontSize: count === 4 ? 6.4 : 7.4,
      color: palette.muted,
      margin: 0,
      breakLine: false,
      valign: "top",
      fit: "shrink",
      paraSpaceAfter: 3,
      lineSpacingMultiple: 1.1,
    });
  }

  addFooter(slide, palette, "MY FEATURE");
  await addDecorationSlide(pptx, palette, data.template);
  await pptx.writeFile({
    fileName: `${sanitizeFileName(data.title)}.pptx`,
    compression: true,
  });
}

export async function createMotivationPowerPoint(
  data: MotivationPptData,
) {
  const { pptx, slide, palette } = await createBase(data.template);
  pptx.title = data.title;
  addHeader(slide, palette, data.title, data.name, "MY MOTIVATION STORY");
  const episodes = data.episodes.slice(0, data.count);
  const graphX = 0.98;
  const graphY = 1.34;
  const graphW = 9.95;
  const graphH = 2.12;
  const xAt = (index: number) =>
    graphX + (index * graphW) / Math.max(episodes.length - 1, 1);
  const yAt = (value: number) => graphY + ((100 - value) / 200) * graphH;

  for (const value of [100, 50, 0, -50, -100]) {
    slide.addShape("line", {
      x: graphX,
      y: yAt(value),
      w: graphW,
      h: 0,
      line: {
        color: palette.muted,
        transparency: value === 0 ? 48 : 78,
        width: value === 0 ? 1.1 : 0.6,
      },
    });
    slide.addText(String(value), {
      x: 0.44,
      y: yAt(value) - 0.08,
      w: 0.38,
      h: 0.15,
      align: "right",
      fontFace: "Arial",
      fontSize: 5.5,
      color: palette.muted,
      margin: 0,
    });
  }

  for (let index = 0; index < episodes.length - 1; index += 1) {
    const current = episodes[index];
    const next = episodes[index + 1];
    slide.addShape("line", {
      x: xAt(index),
      y: yAt(current.motivation),
      w: xAt(index + 1) - xAt(index),
      h: yAt(next.motivation) - yAt(current.motivation),
      line: {
        color: palette.accent,
        width: 3,
        beginArrowType: "none",
        endArrowType: "none",
      },
    });
  }

  for (const [index, episode] of episodes.entries()) {
    const pointX = xAt(index);
    const pointY = yAt(episode.motivation);
    slide.addShape("ellipse", {
      x: pointX - 0.095,
      y: pointY - 0.095,
      w: 0.19,
      h: 0.19,
      line: { color: palette.accent, width: 2 },
      fill: { color: palette.bg },
    });
    slide.addText(String(episode.motivation), {
      x: pointX - 0.3,
      y: pointY - 0.35,
      w: 0.6,
      h: 0.16,
      align: "center",
      fontFace: "Arial",
      fontSize: 6.5,
      bold: true,
      color: palette.fg,
      margin: 0,
    });
    slide.addText(episode.period, {
      x: pointX - 0.52,
      y: graphY + graphH + 0.11,
      w: 1.04,
      h: 0.18,
      align: "center",
      fontFace: "Yu Gothic",
      fontSize: 5.7,
      bold: true,
      color: palette.muted,
      margin: 0,
      fit: "shrink",
    });
  }

  const summaryY = 3.86;
  const summaryGap = 0.2;
  const summaryW = (10.47 - summaryGap) / 2;
  const summaryItems = [
    {
      label: "私がモチベーションが上がる時",
      value: data.motivationUpSummary || "—",
      arrow: "↗",
    },
    {
      label: "私がモチベーションが下がる時",
      value: data.motivationDownSummary || "—",
      arrow: "↘",
    },
  ];

  summaryItems.forEach((item, index) => {
    const x = 0.61 + index * (summaryW + summaryGap);
    slide.addShape("roundRect", {
      x,
      y: summaryY,
      w: summaryW,
      h: 0.72,
      line: { color: palette.muted, transparency: 82, width: 0.6 },
      fill: { color: palette.soft },
    });
    slide.addShape("ellipse", {
      x: x + 0.16,
      y: summaryY + 0.18,
      w: 0.34,
      h: 0.34,
      line: { color: palette.accent, transparency: 100 },
      fill: { color: palette.accent },
    });
    slide.addText(item.arrow, {
      x: x + 0.16,
      y: summaryY + 0.24,
      w: 0.34,
      h: 0.12,
      align: "center",
      fontFace: "Arial",
      fontSize: 7,
      bold: true,
      color: "20201F",
      margin: 0,
    });
    slide.addText(item.label, {
      x: x + 0.62,
      y: summaryY + 0.12,
      w: summaryW - 0.78,
      h: 0.17,
      fontFace: "Yu Gothic",
      fontSize: 6,
      bold: true,
      color: palette.muted,
      margin: 0,
      fit: "shrink",
    });
    slide.addText(item.value, {
      x: x + 0.62,
      y: summaryY + 0.34,
      w: summaryW - 0.78,
      h: 0.25,
      fontFace: "Yu Gothic",
      fontSize: 6.7,
      bold: true,
      color: palette.fg,
      margin: 0,
      fit: "shrink",
      valign: "middle",
    });
  });

  const gridY = 4.83;
  const columns = data.count === 7 ? 4 : 3;
  const rows = Math.ceil(data.count / columns);
  const gapX = 0.18;
  const gapY = 0.17;
  const cardW = (10.47 - gapX * (columns - 1)) / columns;
  const cardH = (2.56 - gapY * (rows - 1)) / rows;

  for (const [index, episode] of episodes.entries()) {
    const column = index % columns;
    const row = Math.floor(index / columns);
    const x = 0.61 + column * (cardW + gapX);
    const y = gridY + row * (cardH + gapY);
    slide.addShape("ellipse", {
      x,
      y,
      w: 0.32,
      h: 0.32,
      line: { color: palette.accent, transparency: 100 },
      fill: { color: palette.accent },
    });
    slide.addText(String(index + 1), {
      x,
      y: y + 0.06,
      w: 0.32,
      h: 0.11,
      align: "center",
      fontFace: "Arial",
      fontSize: 5.5,
      bold: true,
      color: "20201F",
      margin: 0,
    });
    slide.addText(episode.period, {
      x: x + 0.44,
      y,
      w: cardW - 0.44,
      h: 0.16,
      fontFace: "Yu Gothic",
      fontSize: 5.3,
      color: palette.muted,
      margin: 0,
    });
    slide.addText(episode.title, {
      x: x + 0.44,
      y: y + 0.2,
      w: cardW - 0.44,
      h: 0.23,
      fontFace: "Yu Gothic",
      fontSize: 8.2,
      bold: true,
      color: palette.fg,
      margin: 0,
      fit: "shrink",
    });
    slide.addText(episode.description, {
      x: x + 0.44,
      y: y + 0.47,
      w: cardW - 0.44,
      h: Math.max(0.32, cardH - 0.52),
      fontFace: "Yu Gothic",
      fontSize: 5.3,
      color: palette.muted,
      margin: 0,
      fit: "shrink",
      valign: "top",
      lineSpacingMultiple: 1.05,
    });
    if (episode.rookiesLogoId) {
      const logo = await logoDataUri(episode.rookiesLogoId, palette);
      const logoBox = fitLogo(
        logo,
        x + cardW - 0.92,
        y + cardH - 0.26,
        0.85,
        0.18,
      );
      slide.addImage({
        data: logo.data,
        ...logoBox,
      });
    }
  }

  addFooter(slide, palette, "MOTIVATION GRAPH");
  await addDecorationSlide(pptx, palette, data.template);
  await pptx.writeFile({
    fileName: `${sanitizeFileName(data.title)}.pptx`,
    compression: true,
  });
}
