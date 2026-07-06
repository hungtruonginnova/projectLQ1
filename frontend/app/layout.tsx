import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Overdue Parts Report",
  description: "Generate the weekly Seawind overdue-parts report from four ERP exports.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
