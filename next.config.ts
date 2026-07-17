import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  typescript: { ignoreBuildErrors: true },
  webpack: (config, { isServer }) => {
    if (!isServer) {
      config.resolve.fallback = {
        ...config.resolve.fallback,
        https: false, http: false, net: false, tls: false,
        fs: false, crypto: false, stream: false, zlib: false,
        path: false, os: false, url: false,
      }
    }
    return config
  },
};

export default nextConfig;
