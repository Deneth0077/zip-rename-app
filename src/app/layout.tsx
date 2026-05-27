import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "ZIP Rename Nexus | NIC to Employee Code PDF Smart Relabeler",
  description: "Cross-match and rename PDF files from ZIP archives using Excel worksheets containing NIC numbers and Employee Codes. Processes 100% locally.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html
      lang="en"
      className="h-full antialiased"
      suppressHydrationWarning
    >
      <body className="min-h-full flex flex-col font-sans" suppressHydrationWarning>{children}</body>
    </html>
  );
}
