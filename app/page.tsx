import type { Metadata } from "next";
import { WatashiNoKaoApp } from "./watashi-no-kao-app";

export const metadata: Metadata = {
  title: "自己紹介PowerPointをかんたん作成",
  description:
    "質問に答えるだけで、自分らしさが伝わる自己紹介PowerPointを作成できます。",
};

export default function Home() {
  return <WatashiNoKaoApp />;
}
