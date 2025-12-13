// helpers
import { outputAuto } from '@/lib/helpers/outputPNG';

// necessary imports from png collection state slice
import { ImageData } from '@/lib/state/slices/pngCollectionSlice';

// interfaces and types
import { Template } from '@/lib/helpers/outputPNG';
import { ApiError } from '@/app/api/api-Errors';
import { NextRequest, NextResponse } from 'next/server';

export async function POST(req: NextRequest) {
  try {
    //req body: {template, images: [array of image blocks]}
    //image blocks: {url: "url.com", spine_color: "#ffffffff", type}
    const { template, images }: { template: Template; images: ImageData[] } =
      await req.json();
    const { mime, filename, buffer } = await outputAuto({
      template,
      images,
      prefix: 'grid',
    });

    return new NextResponse(new Uint8Array(buffer), {
      status: 201,
      headers: {
        'Content-Type': mime,
        'Content-Disposition': `attachment; filename="${filename}"`,
        'Content-Length': String(buffer.length),
        'Cache-Control': 'no-store',
      },
    });
  } catch {
    return new ApiError(
      400,
      'PNG Creation Error',
      'There was an error during PNG creation, please try again',
    ).format();
  }
}
