import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Excel Gantt Generator",
  description: "백데이터 기반 간트차트 Excel 생성기"
};

export default function RootLayout({
  children
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="ko">
      <body>{children}</body>
    </html>
  );
}
