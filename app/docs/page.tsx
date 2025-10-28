'use client';

import SwaggerUI from 'swagger-ui-react';

export default function DocsPage() {
  return (
    <SwaggerUI
      url="/api/openapi"
      tryItOutEnabled={false}
      supportedSubmitMethods={[]}
      docExpansion="list"
    />
  );
}
