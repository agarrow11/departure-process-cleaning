/** @type {import('next').NextConfig} */
const nextConfig = {
  // Required for SheetJS (xlsx) to work correctly in API routes
  serverExternalPackages: ['xlsx'],
  typescript: {
    ignoreBuildErrors: true,
  },
  images: {
    unoptimized: true,
  },
}

export default nextConfig
