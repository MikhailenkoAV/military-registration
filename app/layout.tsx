import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Воинский учёт — рабочий контур",
  description: "Локальная программа для ведения воинского учёта, контроля сроков и подготовки документов.",
  other: {
    "codex-preview": "development",
  },
  icons: {
    icon: "/favicon.svg",
    shortcut: "/favicon.svg",
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="ru">
      <body>{children}</body>
    </html>
  );
}
