export function readFileAsBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result;
      if (typeof result !== 'string') {
        reject(new Error('Unable to read uploaded file.'));
        return;
      }
      const [, dataBase64 = ''] = result.split(',');
      resolve(dataBase64);
    };
    reader.onerror = () => {
      reject(new Error('Unable to read uploaded file.'));
    };
    reader.readAsDataURL(file);
  });
}
