/** @type {import('next').NextConfig} */
const nextConfig = {
  async headers() {
    const allowedOrigin =
      process.env.NEXT_PUBLIC_SITE_URL ||
      'https://megs-comfort-creations.vercel.app';

    return [
      {
        source: '/api/:path*',
        headers: [
          { key: 'Access-Control-Allow-Origin', value: allowedOrigin },
          { key: 'Access-Control-Allow-Credentials', value: 'true' },
          {
            key: 'Access-Control-Allow-Methods',
            value: 'GET,POST,PUT,PATCH,DELETE,OPTIONS',
          },
          {
            key: 'Access-Control-Allow-Headers',
            value:
              'Origin, X-Requested-With, Content-Type, Accept, Authorization',
          },
        ],
      },
    ];
  },
};

module.exports = nextConfig;
