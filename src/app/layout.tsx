import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "genOferta AI — объединение изменений в Оферту",
  description:
    "Гибридная сборка изменений в Публичную оферту: детерминированный движок правит документ, ИИ помогает разобрать нестандартные формулировки.",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="ru">
      <body>{children}</body>
    </html>
  );
}
