"use client";

import {
  ChangeEvent,
  CSSProperties,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import styles from "./watashi-no-kao.module.css";
import {
  createFacePowerPoint,
  createMotivationPowerPoint,
  type FacePptData,
  type MotivationPptData,
} from "./pptx";

type ProjectType = "face" | "motivation";
type TemplateId = "simple" | "cool" | "fashion";
type LogoId =
  | "rookies"
  | "gp-2024"
  | "gp-2025"
  | "gp-2026"
  | "guild"
  | "world-series"
  | null;

type PhotoData = {
  dataUrl: string;
  name: string;
  zoom: number;
  x: number;
  y: number;
};

type FaceBlock = {
  id: string;
  title: string;
  description: string;
  photo: PhotoData | null;
  rookiesLogoId: LogoId;
};

type Episode = {
  id: string;
  period: string;
  title: string;
  description: string;
  motivation: number;
  rookiesLogoId: LogoId;
};

type FaceState = {
  name: string;
  title: string;
  count: 3 | 4;
  template: TemplateId;
  blocks: FaceBlock[];
};

type MotivationState = {
  name: string;
  title: string;
  count: 5 | 6 | 7;
  template: TemplateId;
  motivationUpSummary: string;
  motivationDownSummary: string;
  episodes: Episode[];
};

type SavedState = {
  project: ProjectType | null;
  face: FaceState;
  motivation: MotivationState;
};

const STORAGE_KEY = "watashi-no-kao-v1";
const MAX_IMAGE_SIZE = 4 * 1024 * 1024;

const logoOptions: Array<{
  id: LogoId;
  title: string;
  short: string;
  image?: string;
}> = [
  { id: null, title: "設定しない", short: "なし" },
  {
    id: "rookies",
    title: "ROOKIES",
    short: "ROOKIES",
    image: "/logos/rookies/rookies.png",
  },
  {
    id: "gp-2024",
    title: "CAREER ROOKIES GP 2024",
    short: "GP '24",
    image: "/logos/rookies/gp-2024.png",
  },
  {
    id: "gp-2025",
    title: "CAREER ROOKIES GP 2025",
    short: "GP '25",
    image: "/logos/rookies/gp-2025.png",
  },
  {
    id: "gp-2026",
    title: "CAREER ROOKIES GP 2026",
    short: "GP '26",
    image: "/logos/rookies/gp-2026.png",
  },
  {
    id: "guild",
    title: "ROOKIES GUILD",
    short: "GUILD",
    image: "/logos/rookies/guild.png",
  },
  {
    id: "world-series",
    title: "ROOKIES WORLD SERIES",
    short: "WORLD SERIES",
    image: "/logos/rookies/world-series.png",
  },
];

const templateOptions: Array<{
  id: TemplateId;
  eyebrow: string;
  title: string;
  description: string;
  colors: string[];
}> = [
  {
    id: "simple",
    eyebrow: "01",
    title: "シンプル",
    description: "白を基調にした、就活・授業で使いやすいデザイン",
    colors: ["#fffef9", "#d8ff4f", "#20201f"],
  },
  {
    id: "cool",
    eyebrow: "02",
    title: "クール",
    description: "ダークカラーで引き締めた、ビジネス向けデザイン",
    colors: ["#1d2538", "#8fa7ff", "#f8f8f4"],
  },
  {
    id: "fashion",
    eyebrow: "03",
    title: "ファッション",
    description: "写真とタイポグラフィを大胆に使う雑誌風デザイン",
    colors: ["#f5eee6", "#ff7548", "#243f37"],
  },
];

const makeFaceBlock = (index: number): FaceBlock => ({
  id: `face-${Date.now()}-${index}`,
  title: ["好奇心で動く", "チームで挑む", "最後までやりきる", "人を笑顔にする"][
    index
  ],
  description: [
    "気になったことは、まず自分で試してみます。新しい環境や考え方との出会いが、私の原動力です。",
    "一人では届かない目標も、仲間の強みをつなぐことで実現できると考えています。",
    "難しい状況でも、小さな前進を積み重ねます。粘り強さには自信があります。",
    "相手の気持ちを想像し、場が前向きになるコミュニケーションを大切にしています。",
  ][index],
  photo: null,
  rookiesLogoId: index === 1 ? "rookies" : null,
});

const makeEpisode = (index: number): Episode => ({
  id: `episode-${Date.now()}-${index}`,
  period: ["中学3年", "高校1年", "高校3年", "大学1年", "大学3年", "現在", "これから"][
    index
  ],
  title: [
    "部活動に熱中",
    "新しい環境",
    "大会で悔しい経験",
    "仲間との出会い",
    "企画に挑戦",
    "成長を実感",
    "次の挑戦へ",
  ][index],
  description: [
    "毎日の練習を通じて、継続することの強さを学びました。",
    "知り合いのいない環境で、自分から話しかけることを意識しました。",
    "目標に届かず、努力の方法を見直すきっかけになりました。",
    "価値観の違う仲間と協力する楽しさを知りました。",
    "学生イベントの運営に挑戦。最後までやり切りました。",
    "経験がつながり、自分の強みを言葉にできるようになりました。",
    "学びを活かし、周りを巻き込む挑戦を続けたいです。",
  ][index],
  motivation: [40, 10, -50, 30, 80, 60, 90][index],
  rookiesLogoId: index === 4 ? "gp-2026" : null,
});

const initialFace: FaceState = {
  name: "山田 花子",
  title: "山田花子を表す3つの特徴",
  count: 3,
  template: "simple",
  blocks: [0, 1, 2, 3].map(makeFaceBlock),
};

const initialMotivation: MotivationState = {
  name: "山田 花子",
  title: "私のモチベーショングラフ",
  count: 6,
  template: "simple",
  motivationUpSummary:
    "目標が明確で、仲間と進捗を共有しながら挑戦できるとき。",
  motivationDownSummary:
    "努力の手応えが見えず、一人で抱え込んでしまうとき。",
  episodes: [0, 1, 2, 3, 4, 5, 6].map(makeEpisode),
};

const logoLabel = (id: LogoId) =>
  logoOptions.find((option) => option.id === id)?.short ?? "";

const logoImage = (id: LogoId) =>
  logoOptions.find((option) => option.id === id)?.image;

const templateClass = (id: TemplateId) =>
  id === "simple"
    ? styles.templateSimple
    : id === "cool"
      ? styles.templateCool
      : styles.templateFashion;

function CharacterCount({
  value,
  max,
}: {
  value: string;
  max: number;
}) {
  const left = max - value.length;
  return (
    <span className={left < 6 ? styles.countDanger : styles.characterCount}>
      あと{left}文字
    </span>
  );
}

function StepHeader({
  step,
  project,
  onStep,
}: {
  step: number;
  project: ProjectType;
  onStep: (value: number) => void;
}) {
  const labels = ["デザイン", "基本情報", "内容入力", "プレビュー", "ダウンロード"];
  return (
    <div className={styles.stepHeader}>
      <div className={styles.projectPill}>
        <span>{project === "face" ? "MY FEATURE" : "モチベーショングラフ"}</span>
        <b>{step + 1}/5</b>
      </div>
      <div className={styles.steps} aria-label="作成ステップ">
        {labels.map((label, index) => (
          <button
            type="button"
            key={label}
            className={`${styles.step} ${step === index ? styles.stepActive : ""} ${
              step > index ? styles.stepDone : ""
            }`}
            onClick={() => onStep(index)}
            aria-current={step === index ? "step" : undefined}
          >
            <span>{step > index ? "✓" : index + 1}</span>
            {label}
          </button>
        ))}
      </div>
    </div>
  );
}

function TemplateSelector({
  value,
  onChange,
}: {
  value: TemplateId;
  onChange: (value: TemplateId) => void;
}) {
  return (
    <div className={styles.templateGrid}>
      {templateOptions.map((template) => (
        <button
          type="button"
          key={template.id}
          className={`${styles.templateCard} ${
            value === template.id ? styles.templateCardActive : ""
          }`}
          onClick={() => onChange(template.id)}
          aria-pressed={value === template.id}
        >
          <span className={styles.templateNumber}>{template.eyebrow}</span>
          <span className={styles.templateSwatches}>
            {template.colors.map((color) => (
              <i key={color} style={{ background: color }} />
            ))}
          </span>
          <strong>{template.title}</strong>
          <small>{template.description}</small>
          <span className={styles.selectedTick}>
            {value === template.id ? "✓ 選択中" : "選択する"}
          </span>
        </button>
      ))}
    </div>
  );
}

function LogoSelector({
  value,
  onChange,
}: {
  value: LogoId;
  onChange: (value: LogoId) => void;
}) {
  return (
    <div className={styles.logoGrid}>
      {logoOptions.map((logo) => (
        <button
          type="button"
          key={logo.id ?? "none"}
          onClick={() => onChange(logo.id)}
          className={`${styles.logoCard} ${
            value === logo.id ? styles.logoCardActive : ""
          }`}
          aria-pressed={value === logo.id}
          title={logo.title}
        >
          <span
            className={`${styles.logoMark} ${
              logo.id === null ? styles.logoNone : ""
            }`}
          >
            {logo.image ? (
              // eslint-disable-next-line @next/next/no-img-element
              <img src={logo.image} alt="" />
            ) : logo.id === null ? (
              "—"
            ) : (
              logo.short
            )}
          </span>
          <small>{logo.id === null ? "設定しない" : logo.title}</small>
        </button>
      ))}
    </div>
  );
}

function PreviewLogo({ id }: { id: Exclude<LogoId, null> }) {
  const image = logoImage(id);
  return image ? (
    <span
      className={`${styles.previewLogo} ${styles.previewLogoImage}`}
      data-logo-id={id}
    >
      {/* eslint-disable-next-line @next/next/no-img-element */}
      <img src={image} alt={logoOptions.find((option) => option.id === id)?.title} />
    </span>
  ) : (
    <span className={styles.previewLogo}>{logoLabel(id)}</span>
  );
}

function PhotoPlaceholder({ index }: { index: number }) {
  return (
    <div className={styles.photoPlaceholder}>
      <span>{String(index + 1).padStart(2, "0")}</span>
      <b>PHOTO</b>
    </div>
  );
}

async function compressImage(file: File): Promise<string> {
  const objectUrl = URL.createObjectURL(file);
  try {
    const image = new Image();
    image.src = objectUrl;
    await image.decode();
    const max = 1600;
    const scale = Math.min(1, max / Math.max(image.width, image.height));
    const canvas = document.createElement("canvas");
    canvas.width = Math.round(image.width * scale);
    canvas.height = Math.round(image.height * scale);
    const context = canvas.getContext("2d");
    if (!context) throw new Error("画像を読み込めませんでした。");
    context.drawImage(image, 0, 0, canvas.width, canvas.height);
    return canvas.toDataURL("image/jpeg", 0.86);
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

function FacePreview({ data }: { data: FaceState }) {
  const blocks = data.blocks.slice(0, data.count);
  return (
    <div className={`${styles.slide} ${templateClass(data.template)}`}>
      <div className={styles.slideAccent} />
      <header className={styles.slideHeader}>
        <div>
          <span>ABOUT ME / {String(data.count).padStart(2, "0")} FACES</span>
          <h2>{data.title || "資料タイトルを入力してください"}</h2>
        </div>
        <p>{data.name || "氏名"}</p>
      </header>
      <div
        className={`${styles.faceGrid} ${
          data.count === 4 ? styles.faceGridFour : ""
        }`}
      >
        {blocks.map((block, index) => (
          <article className={styles.facePreviewCard} key={block.id}>
            <div className={styles.facePhoto}>
              {block.photo ? (
                // eslint-disable-next-line @next/next/no-img-element
                <img
                  src={block.photo.dataUrl}
                  alt=""
                  style={{
                    transform: `scale(${block.photo.zoom})`,
                    objectPosition: `${block.photo.x}% ${block.photo.y}%`,
                  }}
                />
              ) : (
                <PhotoPlaceholder index={index} />
              )}
              <span className={styles.faceIndex}>0{index + 1}</span>
            </div>
            <div className={styles.faceText}>
              {block.rookiesLogoId && (
                <PreviewLogo id={block.rookiesLogoId} />
              )}
              <h3>{block.title || `特徴 ${index + 1}`}</h3>
              <p>{block.description || "ここに説明文が入ります。"}</p>
            </div>
          </article>
        ))}
      </div>
      <footer className={styles.slideFooter}>
        <span>MY FEATURE</span>
        <i />
        <span>MY STORY MAKER</span>
      </footer>
    </div>
  );
}

function MotivationPreview({ data }: { data: MotivationState }) {
  const episodes = data.episodes.slice(0, data.count);
  const width = 900;
  const height = 260;
  const left = 58;
  const right = 24;
  const top = 20;
  const bottom = 40;
  const x = (index: number) =>
    left + (index * (width - left - right)) / Math.max(episodes.length - 1, 1);
  const y = (value: number) =>
    top + ((100 - value) / 200) * (height - top - bottom);
  const points = episodes.map((episode, index) => `${x(index)},${y(episode.motivation)}`).join(" ");
  return (
    <div className={`${styles.slide} ${templateClass(data.template)}`}>
      <div className={styles.slideAccent} />
      <header className={styles.slideHeader}>
        <div>
          <span>MY MOTIVATION STORY</span>
          <h2>{data.title || "モチベーショングラフ"}</h2>
        </div>
        <p>{data.name || "氏名"}</p>
      </header>
      <div className={styles.chartWrap}>
        <svg
          viewBox={`0 0 ${width} ${height}`}
          role="img"
          aria-label="モチベーションの推移"
        >
          {[100, 50, 0, -50, -100].map((value) => (
            <g key={value}>
              <line
                x1={left}
                y1={y(value)}
                x2={width - right}
                y2={y(value)}
                className={value === 0 ? styles.axisStrong : styles.axis}
              />
              <text x={12} y={y(value) + 4}>
                {value}
              </text>
            </g>
          ))}
          <polyline points={points} className={styles.chartLine} />
          {episodes.map((episode, index) => (
            <g key={episode.id}>
              <circle
                cx={x(index)}
                cy={y(episode.motivation)}
                r="8"
                className={styles.chartPoint}
              />
              <text
                x={x(index)}
                y={height - 10}
                textAnchor="middle"
                className={styles.chartPeriod}
              >
                {episode.period || `${index + 1}`}
              </text>
              <text
                x={x(index)}
                y={y(episode.motivation) - 15}
                textAnchor="middle"
                className={styles.chartValue}
              >
                {episode.motivation}
              </text>
            </g>
          ))}
        </svg>
      </div>
      <div className={styles.motivationSummaryPreview}>
        <article>
          <span>↗</span>
          <div>
            <b>私がモチベーションが上がる時</b>
            <p>{data.motivationUpSummary || "—"}</p>
          </div>
        </article>
        <article>
          <span>↘</span>
          <div>
            <b>私がモチベーションが下がる時</b>
            <p>{data.motivationDownSummary || "—"}</p>
          </div>
        </article>
      </div>
      <div
        className={`${styles.episodeGrid} ${
          data.count === 7 ? styles.episodeGridSeven : ""
        }`}
      >
        {episodes.map((episode, index) => (
          <article key={episode.id}>
            <span className={styles.episodeNumber}>{index + 1}</span>
            <div>
              <small>{episode.period || "時期"}</small>
              <h3>{episode.title || `エピソード ${index + 1}`}</h3>
              <p>{episode.description || "ここに出来事の説明が入ります。"}</p>
              {episode.rookiesLogoId && (
                <PreviewLogo id={episode.rookiesLogoId} />
              )}
            </div>
          </article>
        ))}
      </div>
      <footer className={styles.slideFooter}>
        <span>MOTIVATION GRAPH</span>
        <i />
        <span>MY STORY MAKER</span>
      </footer>
    </div>
  );
}

function InputField({
  label,
  value,
  max,
  placeholder,
  onChange,
}: {
  label: string;
  value: string;
  max: number;
  placeholder?: string;
  onChange: (value: string) => void;
}) {
  return (
    <label className={styles.field}>
      <span>
        <b>{label}</b>
        <CharacterCount value={value} max={max} />
      </span>
      <input
        value={value}
        maxLength={max}
        placeholder={placeholder}
        onChange={(event) => onChange(event.target.value)}
      />
    </label>
  );
}

function TextAreaField({
  label,
  value,
  max,
  placeholder,
  onChange,
}: {
  label: string;
  value: string;
  max: number;
  placeholder?: string;
  onChange: (value: string) => void;
}) {
  return (
    <label className={styles.field}>
      <span>
        <b>{label}</b>
        <CharacterCount value={value} max={max} />
      </span>
      <textarea
        value={value}
        maxLength={max}
        placeholder={placeholder}
        onChange={(event) => onChange(event.target.value)}
      />
    </label>
  );
}

function Welcome({
  onSelect,
  onRestore,
  hasSaved,
}: {
  onSelect: (project: ProjectType) => void;
  onRestore: () => void;
  hasSaved: boolean;
}) {
  return (
    <main className={styles.welcome}>
      <nav className={styles.welcomeNav}>
        <div className={styles.brand}>
          <span className={styles.brandFace}>☺</span>
          <div>
            <b>MY STORY MAKER</b>
            <small>SELF DISCOVERY POWERPOINT</small>
          </div>
        </div>
        <span className={styles.pcNote}>PCでの利用を推奨しています</span>
      </nav>

      <section className={styles.hero}>
        <div className={styles.heroCopy}>
          <span className={styles.kicker}>POWERPOINT SELF-PORTRAIT TOOL</span>
          <h1>
            自分を伝える
            <br />
            PowerPointを、
            <br />
            <em>かんたんにつくろう。</em>
          </h1>
          <p>
            質問に答えていくだけで、あなたの経験や強みが
            <br />
            伝わる自己紹介資料ができあがります。
          </p>
          <div className={styles.heroMeta}>
            <span>01 振り返る</span>
            <i />
            <span>02 言葉にする</span>
            <i />
            <span>03 伝える</span>
          </div>
        </div>
        <div className={styles.heroVisual} aria-hidden="true">
          <div className={styles.heroBlob}>
            <span>ME</span>
            <b>私らしさを、<br />一枚に。</b>
          </div>
          <span className={styles.heroDotOne}>●</span>
          <span className={styles.heroDotTwo}>✦</span>
          <span className={styles.heroLine}>POWERPOINT / EDITABLE / A4</span>
        </div>
      </section>

      <section className={styles.chooseSection}>
        <div className={styles.sectionHeading}>
          <span>STEP 01</span>
          <div>
            <h2>作成する資料を選んでください</h2>
            <p>あとからトップ画面に戻って変更できます</p>
          </div>
        </div>
        <div className={styles.projectCards}>
          <button type="button" onClick={() => onSelect("face")}>
            <span className={styles.projectTag}>SELF INTRODUCTION</span>
            <div className={styles.faceMiniPreview}>
              {[0, 1, 2].map((index) => (
                <i key={index}>
                  <span>0{index + 1}</span>
                </i>
              ))}
            </div>
            <div>
              <small>01</small>
              <h3>MY FEATURE</h3>
              <p>自分をつくる3つ・4つの要素を、一枚で紹介</p>
              <b>この資料をつくる <span>→</span></b>
            </div>
          </button>
          <button type="button" onClick={() => onSelect("motivation")}>
            <span className={styles.projectTag}>MY STORY</span>
            <div className={styles.graphMiniPreview}>
              <i />
              {[24, 60, 35, 76, 55, 86].map((top, index) => (
                <span
                  key={index}
                  style={{ "--point-top": `${100 - top}%` } as CSSProperties}
                />
              ))}
            </div>
            <div>
              <small>02</small>
              <h3>モチベーショングラフ</h3>
              <p>人生の出来事と気持ちの変化を、線で可視化</p>
              <b>この資料をつくる <span>→</span></b>
            </div>
          </button>
        </div>
        {hasSaved && (
          <button type="button" className={styles.restoreButton} onClick={onRestore}>
            <span>↻</span> 前回の続きから再開する
          </button>
        )}
      </section>

      <footer className={styles.welcomeFooter}>
        <span>入力内容はこのブラウザに自動保存されます</span>
          <b>Copyright © 2026 Strobolights All Rights Reserved.</b>
      </footer>
    </main>
  );
}

export function WatashiNoKaoApp() {
  const [project, setProject] = useState<ProjectType | null>(null);
  const [face, setFace] = useState<FaceState>(initialFace);
  const [motivation, setMotivation] =
    useState<MotivationState>(initialMotivation);
  const [step, setStep] = useState(0);
  const [hydrated, setHydrated] = useState(false);
  const [savedProject, setSavedProject] = useState<ProjectType | null>(null);
  const [savedAt, setSavedAt] = useState<Date | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [downloading, setDownloading] = useState(false);
  const [openItem, setOpenItem] = useState(0);
  const saveTimer = useRef<ReturnType<typeof setTimeout> | null>(null);

  useEffect(() => {
    let cancelled = false;
    let saved: SavedState | null = null;
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        saved = JSON.parse(raw) as SavedState;
      }
    } catch {
      localStorage.removeItem(STORAGE_KEY);
    }
    queueMicrotask(() => {
      if (cancelled) return;
      if (saved) {
        setSavedProject(saved.project);
        if (saved.face) {
          const compactName = saved.face.name.replace(/\s/g, "");
          const oldDefaultTitle = `${compactName}を表す${saved.face.count}つの顔`;
          setFace({
            ...saved.face,
            title:
              saved.face.title === oldDefaultTitle
                ? `${compactName}を表す${saved.face.count}つの特徴`
                : saved.face.title,
          });
        }
        if (saved.motivation) {
          setMotivation({
            ...initialMotivation,
            ...saved.motivation,
            motivationUpSummary:
              saved.motivation.motivationUpSummary ?? "",
            motivationDownSummary:
              saved.motivation.motivationDownSummary ?? "",
            episodes:
              saved.motivation.episodes ?? initialMotivation.episodes,
          });
        }
      }
      setHydrated(true);
    });
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    if (!hydrated) return;
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => {
      try {
        const state: SavedState = { project, face, motivation };
        localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
        setSavedProject(project);
        setSavedAt(new Date());
      } catch {
        setError(
          "自動保存の容量がいっぱいです。写真の枚数を減らすか、小さい画像に差し替えてください。",
        );
      }
    }, 450);
    return () => {
      if (saveTimer.current) clearTimeout(saveTimer.current);
    };
  }, [hydrated, project, face, motivation]);

  const activeTemplate = project === "face" ? face.template : motivation.template;
  const activeName = project === "face" ? face.name : motivation.name;
  const activeTitle = project === "face" ? face.title : motivation.title;

  const validationIssues = useMemo(() => {
    const issues: string[] = [];
    if (!activeName.trim()) issues.push("氏名");
    if (!activeTitle.trim()) issues.push("資料タイトル");
    if (project === "face") {
      face.blocks.slice(0, face.count).forEach((block, index) => {
        if (!block.title.trim()) issues.push(`特徴 ${index + 1}のタイトル`);
        if (!block.description.trim()) issues.push(`特徴 ${index + 1}の説明文`);
      });
    }
    if (project === "motivation") {
      motivation.episodes
        .slice(0, motivation.count)
        .forEach((episode, index) => {
          if (!episode.period.trim()) issues.push(`エピソード ${index + 1}の時期`);
          if (!episode.title.trim())
            issues.push(`エピソード ${index + 1}のタイトル`);
          if (!episode.description.trim())
            issues.push(`エピソード ${index + 1}の説明文`);
        });
    }
    return issues;
  }, [project, activeName, activeTitle, face, motivation]);

  const selectProject = (selected: ProjectType) => {
    setProject(selected);
    setStep(0);
    setOpenItem(0);
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  const updateFaceBlock = (
    index: number,
    patch: Partial<FaceBlock>,
  ) => {
    setFace((current) => ({
      ...current,
      blocks: current.blocks.map((block, blockIndex) =>
        blockIndex === index ? { ...block, ...patch } : block,
      ),
    }));
  };

  const updateEpisode = (index: number, patch: Partial<Episode>) => {
    setMotivation((current) => ({
      ...current,
      episodes: current.episodes.map((episode, episodeIndex) =>
        episodeIndex === index ? { ...episode, ...patch } : episode,
      ),
    }));
  };

  const handleImage = async (
    event: ChangeEvent<HTMLInputElement>,
    index: number,
  ) => {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    if (!["image/jpeg", "image/png", "image/webp"].includes(file.type)) {
      setError("JPEG、PNG、WebP形式の画像を選択してください。");
      return;
    }
    if (file.size > MAX_IMAGE_SIZE) {
      setError("画像は1枚4MB以内にしてください。");
      return;
    }
    try {
      const dataUrl = await compressImage(file);
      updateFaceBlock(index, {
        photo: { dataUrl, name: file.name, zoom: 1, x: 50, y: 50 },
      });
    } catch {
      setError("画像を読み込めませんでした。別の画像をお試しください。");
    }
  };

  const downloadPptx = async () => {
    if (!project) return;
    if (validationIssues.length) {
      setError(
        `未入力の項目があります：${validationIssues.slice(0, 3).join("、")}${
          validationIssues.length > 3 ? " など" : ""
        }`,
      );
      return;
    }
    setDownloading(true);
    try {
      if (project === "face") {
        await createFacePowerPoint(face as FacePptData);
      } else {
        await createMotivationPowerPoint(motivation as MotivationPptData);
      }
    } catch (downloadError) {
      console.error(downloadError);
      setError(
        "PowerPointの生成に失敗しました。入力内容は保存されています。少し時間をおいてもう一度お試しください。",
      );
    } finally {
      setDownloading(false);
    }
  };

  if (!hydrated) {
    return (
      <div className={styles.loading}>
        <span className={styles.brandFace}>☺</span>
        <b>あなたらしさを準備しています...</b>
      </div>
    );
  }

  if (!project) {
    return (
      <Welcome
        onSelect={selectProject}
        hasSaved={Boolean(savedProject)}
        onRestore={() => {
          if (savedProject) selectProject(savedProject);
        }}
      />
    );
  }

  const renderPanel = () => {
    if (step === 0) {
      return (
        <>
          <div className={styles.panelIntro}>
            <span>DESIGN TEMPLATE</span>
            <h1>どんな印象で伝えたいですか？</h1>
            <p>
              目的や自分らしさに合うデザインを選びましょう。
              いつでも変更できます。
            </p>
          </div>
          <TemplateSelector
            value={activeTemplate}
            onChange={(template) =>
              project === "face"
                ? setFace((current) => ({ ...current, template }))
                : setMotivation((current) => ({ ...current, template }))
            }
          />
        </>
      );
    }

    if (step === 1) {
      return (
        <>
          <div className={styles.panelIntro}>
            <span>BASIC INFORMATION</span>
            <h1>まず、あなたのことを教えてください</h1>
            <p>入力した内容は資料の上部に大きく表示されます。</p>
          </div>
          <div className={styles.formStack}>
            <InputField
              label="氏名"
              value={activeName}
              max={30}
              placeholder="例：山田 花子"
              onChange={(name) =>
                project === "face"
                  ? setFace((current) => ({ ...current, name }))
                  : setMotivation((current) => ({ ...current, name }))
              }
            />
            <InputField
              label="資料タイトル"
              value={activeTitle}
              max={30}
              placeholder={
                project === "face"
                  ? "例：山田花子を表す3つの特徴"
                  : "例：私のモチベーショングラフ"
              }
              onChange={(title) =>
                project === "face"
                  ? setFace((current) => ({ ...current, title }))
                  : setMotivation((current) => ({ ...current, title }))
              }
            />
            <fieldset className={styles.segmentField}>
              <legend>
                {project === "face" ? "特徴の数" : "エピソード数"}
              </legend>
              <div>
                {(project === "face" ? [3, 4] : [5, 6, 7]).map((count) => (
                  <button
                    type="button"
                    key={count}
                    className={
                      (project === "face" ? face.count : motivation.count) === count
                        ? styles.segmentActive
                        : ""
                    }
                    onClick={() => {
                      if (project === "face") {
                        setFace((current) => {
                          const next = count as 3 | 4;
                          const name = current.name.replace(/\s/g, "");
                          return {
                            ...current,
                            count: next,
                            title:
                              current.title ===
                                `${name}を表す${current.count}つの特徴` ||
                              current.title === ""
                                ? `${name}を表す${next}つの特徴`
                                : current.title,
                          };
                        });
                      } else {
                        setMotivation((current) => ({
                          ...current,
                          count: count as 5 | 6 | 7,
                        }));
                      }
                    }}
                  >
                    {count}
                    <small>{project === "face" ? "つ" : "個"}</small>
                  </button>
                ))}
              </div>
            </fieldset>
            <aside className={styles.tip}>
              <span>Q</span>
              <div>
                <b>タイトルに迷ったら</b>
                <p>
                  「自分の名前＋伝えたいテーマ」にすると、初めて見る人にも内容が伝わります。
                </p>
              </div>
            </aside>
          </div>
        </>
      );
    }

    if (step === 2 && project === "face") {
      return (
        <>
          <div className={styles.panelIntro}>
            <span>YOUR FACES</span>
            <h1>あなたをつくる要素は何ですか？</h1>
            <p>経験・価値観・得意なことなど、違う角度から選んでみましょう。</p>
          </div>
          <div className={styles.itemTabs} role="tablist">
            {face.blocks.slice(0, face.count).map((block, index) => (
              <button
                type="button"
                key={block.id}
                role="tab"
                aria-selected={openItem === index}
                className={openItem === index ? styles.itemTabActive : ""}
                onClick={() => setOpenItem(index)}
              >
                <span>0{index + 1}</span>
                {block.title || `特徴 ${index + 1}`}
              </button>
            ))}
          </div>
          {face.blocks.slice(0, face.count).map((block, index) =>
            index === openItem ? (
              <div className={styles.itemEditor} key={block.id}>
                <aside className={styles.questionPrompt}>
                  <span>THINK</span>
                  <p>
                    あなたらしさが最も表れている経験は？
                    <br />
                    その経験から何を学びましたか？
                  </p>
                </aside>
                <InputField
                  label={`特徴 ${index + 1}のタイトル`}
                  value={block.title}
                  max={15}
                  placeholder="例：好奇心で動く"
                  onChange={(title) => updateFaceBlock(index, { title })}
                />
                <TextAreaField
                  label="説明文"
                  value={block.description}
                  max={100}
                  placeholder="具体的な経験や、大切にしている考えを書いてみましょう"
                  onChange={(description) =>
                    updateFaceBlock(index, { description })
                  }
                />
                <div className={styles.photoEditor}>
                  <div className={styles.fieldLabel}>
                    <b>写真</b>
                    <span>JPEG / PNG / WebP・4MBまで</span>
                  </div>
                  <label
                    className={`${styles.uploadArea} ${
                      block.photo ? styles.uploadAreaFilled : ""
                    }`}
                  >
                    {block.photo ? (
                      <>
                        {/* eslint-disable-next-line @next/next/no-img-element */}
                        <img
                          src={block.photo.dataUrl}
                          alt="アップロードした写真"
                          style={{
                            transform: `scale(${block.photo.zoom})`,
                            objectPosition: `${block.photo.x}% ${block.photo.y}%`,
                          }}
                        />
                        <span>写真を差し替える</span>
                      </>
                    ) : (
                      <>
                        <b>＋</b>
                        <span>クリックして写真を選ぶ</span>
                        <small>写真はあとから差し替えられます</small>
                      </>
                    )}
                    <input
                      type="file"
                      accept="image/jpeg,image/png,image/webp"
                      onChange={(event) => handleImage(event, index)}
                    />
                  </label>
                  {block.photo && (
                    <div className={styles.cropControls}>
                      <label>
                        拡大
                        <input
                          type="range"
                          min="1"
                          max="2"
                          step="0.05"
                          value={block.photo.zoom}
                          onChange={(event) =>
                            updateFaceBlock(index, {
                              photo: {
                                ...block.photo!,
                                zoom: Number(event.target.value),
                              },
                            })
                          }
                        />
                      </label>
                      <label>
                        左右
                        <input
                          type="range"
                          min="0"
                          max="100"
                          value={block.photo.x}
                          onChange={(event) =>
                            updateFaceBlock(index, {
                              photo: {
                                ...block.photo!,
                                x: Number(event.target.value),
                              },
                            })
                          }
                        />
                      </label>
                      <label>
                        上下
                        <input
                          type="range"
                          min="0"
                          max="100"
                          value={block.photo.y}
                          onChange={(event) =>
                            updateFaceBlock(index, {
                              photo: {
                                ...block.photo!,
                                y: Number(event.target.value),
                              },
                            })
                          }
                        />
                      </label>
                      <button
                        type="button"
                        onClick={() => updateFaceBlock(index, { photo: null })}
                      >
                        削除
                      </button>
                    </div>
                  )}
                </div>
                <div>
                  <div className={styles.fieldLabel}>
                    <b>ROOKIESロゴ</b>
                    <span>この要素に関連するロゴを選択</span>
                  </div>
                  <LogoSelector
                    value={block.rookiesLogoId}
                    onChange={(rookiesLogoId) =>
                      updateFaceBlock(index, { rookiesLogoId })
                    }
                  />
                </div>
              </div>
            ) : null,
          )}
        </>
      );
    }

    if (step === 2 && project === "motivation") {
      const episode = motivation.episodes[openItem];
      return (
        <>
          <div className={styles.panelIntro}>
            <span>YOUR STORY</span>
            <h1>気持ちが動いた出来事を並べましょう</h1>
            <p>良い出来事だけでなく、悔しさや迷いも成長を伝える材料です。</p>
          </div>
          <div className={styles.itemTabs} role="tablist">
            {motivation.episodes
              .slice(0, motivation.count)
              .map((item, index) => (
                <button
                  type="button"
                  key={item.id}
                  role="tab"
                  aria-selected={openItem === index}
                  className={openItem === index ? styles.itemTabActive : ""}
                  onClick={() => setOpenItem(index)}
                >
                  <span>{index + 1}</span>
                  {item.period || `時期 ${index + 1}`}
                </button>
              ))}
          </div>
          <div className={styles.itemEditor}>
            <aside className={styles.questionPrompt}>
              <span>THINK</span>
              <p>気持ちが上がった・下がったきっかけは何でしたか？</p>
            </aside>
            <InputField
              label="時期"
              value={episode.period}
              max={10}
              placeholder="例：大学3年"
              onChange={(period) => updateEpisode(openItem, { period })}
            />
            <InputField
              label="タイトル"
              value={episode.title}
              max={15}
              placeholder="例：イベント運営に挑戦"
              onChange={(title) => updateEpisode(openItem, { title })}
            />
            <TextAreaField
              label="説明文"
              value={episode.description}
              max={50}
              placeholder="何が起こり、どう感じたかを書きましょう"
              onChange={(description) =>
                updateEpisode(openItem, { description })
              }
            />
            <div className={styles.motivationField}>
              <div className={styles.fieldLabel}>
                <b>モチベーション値</b>
                <span>-100〜100（10刻み）</span>
              </div>
              <output
                className={
                  episode.motivation >= 0
                    ? styles.motivationPositive
                    : styles.motivationNegative
                }
              >
                {episode.motivation > 0 ? "+" : ""}
                {episode.motivation}
              </output>
              <input
                type="range"
                min="-100"
                max="100"
                step="10"
                value={episode.motivation}
                style={
                  {
                    "--range-progress": `${(episode.motivation + 100) / 2}%`,
                  } as CSSProperties
                }
                onChange={(event) =>
                  updateEpisode(openItem, {
                    motivation: Number(event.target.value),
                  })
                }
              />
              <select
                aria-label="モチベーション値を選択"
                value={episode.motivation}
                onChange={(event) =>
                  updateEpisode(openItem, {
                    motivation: Number(event.target.value),
                  })
                }
              >
                {Array.from({ length: 21 }, (_, index) => -100 + index * 10).map(
                  (value) => (
                    <option value={value} key={value}>
                      {value > 0 ? "+" : ""}
                      {value}
                    </option>
                  ),
                )}
              </select>
            </div>
            <div>
              <div className={styles.fieldLabel}>
                <b>ROOKIESロゴ</b>
                <span>この出来事に関連するロゴを選択</span>
              </div>
              <LogoSelector
                value={episode.rookiesLogoId}
                onChange={(rookiesLogoId) =>
                  updateEpisode(openItem, { rookiesLogoId })
                }
              />
            </div>
            <section className={styles.summaryEditor}>
              <div className={styles.summaryHeading}>
                <span>SUMMARY</span>
                <div>
                  <h2>モチベーションの傾向をまとめる</h2>
                  <p>
                    グラフを振り返り、自分の気持ちが動く条件を言葉にしてみましょう。
                  </p>
                </div>
              </div>
              <TextAreaField
                label="私がモチベーションが上がる時"
                value={motivation.motivationUpSummary}
                max={80}
                placeholder="例：目標が明確で、仲間と一緒に挑戦できるとき"
                onChange={(motivationUpSummary) =>
                  setMotivation((current) => ({
                    ...current,
                    motivationUpSummary,
                  }))
                }
              />
              <TextAreaField
                label="私がモチベーションが下がる時"
                value={motivation.motivationDownSummary}
                max={80}
                placeholder="例：努力の手応えが見えず、一人で抱え込むとき"
                onChange={(motivationDownSummary) =>
                  setMotivation((current) => ({
                    ...current,
                    motivationDownSummary,
                  }))
                }
              />
            </section>
          </div>
        </>
      );
    }

    if (step === 3) {
      return (
        <>
          <div className={styles.panelIntro}>
            <span>FINAL CHECK</span>
            <h1>伝えたい内容になっていますか？</h1>
            <p>
              読む人の視点で、タイトルと説明が自然につながっているか確認しましょう。
            </p>
          </div>
          <div className={styles.checkList}>
            <div className={styles.scoreRing}>
              <strong>{validationIssues.length ? "要確認" : "完成"}</strong>
              <span>{validationIssues.length ? `${validationIssues.length}項目` : "100%"}</span>
            </div>
            <div>
              <h3>
                {validationIssues.length
                  ? "入力を確認してください"
                  : "ダウンロードの準備ができました"}
              </h3>
              {validationIssues.length ? (
                <ul>
                  {validationIssues.slice(0, 6).map((issue) => (
                    <li key={issue}>{issue}</li>
                  ))}
                </ul>
              ) : (
                <p>
                  すべての必須項目が入力されています。右のプレビューで文字量や写真の見え方を確認してください。
                </p>
              )}
            </div>
          </div>
          <aside className={styles.tip}>
            <span>TIP</span>
            <div>
              <b>PowerPointで自由に仕上げられます</b>
              <p>
                ダウンロード後も、文字・写真・図形・ロゴを移動・編集できます。
              </p>
            </div>
          </aside>
        </>
      );
    }

    return (
      <>
        <div className={styles.panelIntro}>
          <span>DOWNLOAD</span>
          <h1>あなたらしさを、届けましょう</h1>
          <p>完成した資料をPowerPoint形式で保存します。</p>
        </div>
        <div className={styles.downloadCard}>
          <span className={styles.pptIcon}>P</span>
          <div>
            <h3>{activeTitle || "自己紹介資料"}.pptx</h3>
            <p>A4横・1ページ・編集可能</p>
          </div>
          <button
            type="button"
            onClick={downloadPptx}
            disabled={downloading}
          >
            {downloading ? "生成しています..." : "PowerPointをダウンロード"}
            <span>↓</span>
          </button>
        </div>
        <div className={styles.editableList}>
          <span>✓ テキスト編集可能</span>
          <span>✓ 写真編集可能</span>
          <span>✓ 図形・ロゴ移動可能</span>
        </div>
        {validationIssues.length > 0 && (
          <div className={styles.warningBox}>
            <b>未入力の項目が{validationIssues.length}件あります</b>
            <p>プレビューに戻り、内容を確認してからダウンロードしてください。</p>
            <button type="button" onClick={() => setStep(3)}>
              確認画面へ戻る
            </button>
          </div>
        )}
      </>
    );
  };

  return (
    <div className={styles.appShell}>
      <header className={styles.appTopbar}>
        <button
          type="button"
          className={styles.brandButton}
          onClick={() => setProject(null)}
          aria-label="トップ画面へ戻る"
        >
          <span className={styles.brandFace}>☺</span>
          <div>
            <b>MY STORY MAKER</b>
            <small>SELF DISCOVERY POWERPOINT</small>
          </div>
        </button>
        <div className={styles.saveStatus}>
          <span className={savedAt ? styles.savedDot : styles.savingDot} />
          {savedAt
            ? `${savedAt.toLocaleTimeString("ja-JP", {
                hour: "2-digit",
                minute: "2-digit",
              })} 保存済み`
            : "自動保存中"}
        </div>
        <button
          type="button"
          className={styles.changeProject}
          onClick={() => setProject(null)}
        >
          資料を変更
        </button>
      </header>

      <StepHeader
        step={step}
        project={project}
        onStep={(next) => {
          setStep(next);
          window.scrollTo({ top: 0, behavior: "smooth" });
        }}
      />

      <main className={styles.editorLayout}>
        <section className={styles.formPanel}>
          <div className={styles.formContent}>{renderPanel()}</div>
          <footer className={styles.formNavigation}>
            <button
              type="button"
              className={styles.backButton}
              onClick={() => {
                if (step === 0) setProject(null);
                else setStep((current) => current - 1);
              }}
            >
              ← {step === 0 ? "資料選択へ" : "戻る"}
            </button>
            {step < 4 && (
              <button
                type="button"
                className={styles.nextButton}
                onClick={() => {
                  setStep((current) => Math.min(4, current + 1));
                  window.scrollTo({ top: 0, behavior: "smooth" });
                }}
              >
                次へ進む <span>→</span>
              </button>
            )}
          </footer>
        </section>

        <aside className={styles.previewPanel}>
          <div className={styles.previewTop}>
            <div>
              <span>LIVE PREVIEW</span>
              <b>リアルタイムプレビュー</b>
            </div>
            <span className={styles.a4Badge}>A4 横</span>
          </div>
          <div className={styles.previewStage}>
            <div className={styles.slideScaler}>
              {project === "face" ? (
                <FacePreview data={face} />
              ) : (
                <MotivationPreview data={motivation} />
              )}
            </div>
          </div>
          <p className={styles.previewNote}>
            入力した内容はすぐに反映されます。PowerPointでも同じレイアウトで出力されます。
          </p>
        </aside>
      </main>

      {error && (
        <div className={styles.dialogBackdrop} role="presentation">
          <div className={styles.dialog} role="alertdialog" aria-modal="true">
            <span>!</span>
            <div>
              <h2>確認してください</h2>
              <p>{error}</p>
              <button type="button" onClick={() => setError(null)} autoFocus>
                閉じる
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
