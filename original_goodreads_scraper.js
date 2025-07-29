const axios = require('axios');
const readlineSync = require('readline-sync');
const fs = require('fs');
const path = require('path');

async function fetchGoodreadsData(url) {
    try {
        const response = await axios.get(url);
        return response.data;
    } catch (error) {
        console.error('Error fetching the URL:', error);
        process.exit(1);
    }
}

function parseGoodreadsData(html) {
    const data = [];
    const lines = html.split('\n');
    const searchString = '</div></td>  <td class="field position" style="display: none">';

    for (let i = 0; i < lines.length; i++) {
        if (lines[i].includes(searchString)) {
            if (i + 1 < lines.length) {
                const nextLine = lines[i + 1];
                const match = nextLine.match(/"(.*?)"/);
                if (match && match[1]) {
                    data.push(match[1]);
                }
            }
        }
    }

    return data;
}

async function downloadImage(url, folderPath, index) {
    try {
        const response = await axios.get(url, { responseType: 'arraybuffer' });
        const imagePath = path.join(folderPath, `image${index}.jpg`);
        fs.writeFileSync(imagePath, response.data);
        console.log(`Image saved to ${imagePath}`);
    } catch (error) {
        console.error(`Error downloading image from ${url}:`, error);
    }
}

async function fetchAndDownloadImages(urls) {
    const folderPath = path.join(__dirname, 'Goodreads Book Covers');
    if (!fs.existsSync(folderPath)) {
        fs.mkdirSync(folderPath);
    }

    const downloadPromises = urls.map(async (url, index) => {
        const fullUrl = `https://www.goodreads.com${url}`;
        try {
            const html = await fetchGoodreadsData(fullUrl);
            const lines = html.split('\n');
            const searchString = 'role="presentation"';

            for (let i = 0; i < lines.length; i++) {
                if (lines[i].includes(searchString)) {
                    if (i + 1 < lines.length) {
                        const nextLine = lines[i + 1];
                        const match = nextLine.match(/"(.*?)"/);
                        if (match && match[1]) {
                            const imageUrl = match[1];
                            await downloadImage(imageUrl, folderPath, index);
                        }
                    }
                    break;
                }
            }
        } catch (error) {
            console.error(`Error fetching data from ${fullUrl}:`, error);
        }
    });

    await Promise.all(downloadPromises);
}

async function main() {
    const url = readlineSync.question('Enter the Goodreads URL: ');
    const html = await fetchGoodreadsData(url);
    const data = parseGoodreadsData(html);

    await fetchAndDownloadImages(data);
}

main();
