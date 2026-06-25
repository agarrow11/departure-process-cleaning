/** @type {import('next').NextConfig} */
const nextConfig = {
  // Required for SheetJS (xlsx) and ExcelJS to work correctly in API routes
  serverExternalPackages: ['xlsx', 'exceljs'],
  typescript: {
    ignoreBuildErrors: true,
  },
  images: {
    unoptimized: true,
  },
}

export default nextConfig
