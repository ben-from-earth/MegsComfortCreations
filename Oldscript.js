const form = document.querySelector("form");
const titleInput = document.querySelector("#titleInput");
const authorInput = document.querySelector("#authorInput");
const imgCollection = document.querySelector("#imgCollection");

form.addEventListener("submit", (e) => {
  e.preventDefault();
  let title = titleInput.value;
  let author = authorInput.value;
  title = title.trim().replace(/\s+/g, "+");
  author = author.trim().replace(/\s+/g, "+");
  getCoverImg(title, author);
});

async function getCoverImg(title, author) {
  const searchUrl = `https://www.googleapis.com/books/v1/volumes?q=intitle:${title}+inauthor:${author}`;
  const data = await axios.get(searchUrl);
  for (let item of data.data.items) {
    printInfo(item);
    console.log(item);
  }
}

function printInfo(data) {
  const imgInfo = document.createElement("div");
  imgInfo.classList = "imgInfoBox";
  const img = document.createElement("img");

  const titleLabel = document.createElement("h2");
  titleLabel.innerText = data.volumeInfo.title;

  const authorLabel = document.createElement("h3");
  authorLabel.innerText = "Author: " + data.volumeInfo.authors[0];

  const pubLabel = document.createElement("h3");
  pubDate = data.volumeInfo.publishedDate;
  pubDateItems = pubDate.split("-");
  pubLabel.innerText =
    pubDateItems[1] + "/" + pubDateItems[2] + "/" + pubDateItems[0];

  const pageCount = document.createElement("h3");
  pageCount.innerText = "Page Count: " + data.volumeInfo.pageCount;

  if (data.volumeInfo.imageLinks) {
    let coverURL = data.volumeInfo.imageLinks.thumbnail;
    img.src = coverURL;
  }
  if (parseInt(data.volumeInfo.pageCount) > 0 && img.src) {
    imgInfo.append(titleLabel, img, authorLabel, pubLabel, pageCount);
    imgCollection.appendChild(imgInfo);
  }
}
