import type { Metadata } from "next";
import { headers } from "next/headers";
import "./globals.css";

export async function generateMetadata(): Promise<Metadata> {
  const requestHeaders = await headers();
  const host = requestHeaders.get("host") ?? "localhost:3000";
  const protocol =
    requestHeaders.get("x-forwarded-proto") ??
    (host.includes("localhost") ? "http" : "https");
  const origin = `${protocol}://${host}`;
  const title = "MY STORY MAKER｜自己紹介PowerPointをかんたん作成";
  const description =
    "自分を振り返り、経験を整理し、伝わる自己紹介PowerPointをつくるサービス。";

  return {
    metadataBase: new URL(origin),
    title: {
      default: title,
      template: "%s｜MY STORY MAKER",
    },
    description,
    icons: {
      icon: "/favicon-my-story-maker.png",
      shortcut: "/favicon-my-story-maker.png",
    },
    openGraph: {
      title,
      description,
      type: "website",
      url: origin,
      siteName: "MY STORY MAKER",
      locale: "ja_JP",
      images: [
        {
          url: `${origin}/og-my-story-maker.png`,
          width: 1536,
          height: 1024,
          alt: "MY STORY MAKER — 自分を伝えるPowerPointを、かんたんにつくろう。",
        },
      ],
    },
    twitter: {
      card: "summary_large_image",
      title,
      description,
      images: [`${origin}/og-my-story-maker.png`],
    },
  };
}

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="ja">
      <body>{children}</body>
    </html>
  );
}
