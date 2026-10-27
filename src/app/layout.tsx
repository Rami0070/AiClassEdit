import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "스마트 학급편성 시스템",
  description: "동일 이름 절대 방지, 특수아동 2명 가중치, 스네이크 성적 분산, 남녀 성비 균등화를 완벽 준수하는 학급편성 웹 애플리케이션",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="ko">
      <body className="antialiased bg-slate-50 text-slate-800 selection:bg-emerald-200 selection:text-emerald-900">
        {children}
      </body>
    </html>
  );
}
