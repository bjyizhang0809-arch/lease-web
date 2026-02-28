import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // Allow large file uploads
  api: {
    bodyParser: {
      sizeLimit: "10mb",
    },
  },
};

export default nextConfig;
