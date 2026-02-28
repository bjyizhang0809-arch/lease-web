import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "租赁合同计算器",
  description: "上传租赁合同数据，自动计算应收总额、收入总额及银行发票对账",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="zh-CN">
      <body className="min-h-screen bg-gray-50">{children}</body>
    </html>
  );
}
