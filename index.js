const puppeteer = require('puppeteer');
// const {LevenshteinDistance} = require('./similarity.js');

// load data
var XLSX = require('xlsx');
var workbook = XLSX.readFile("dataIO/Input.xlsx");
var worksheet = workbook.Sheets["Sheet1"];
var data = XLSX.utils.sheet_to_json(worksheet);

const LevenshteinDistance =  (a, b) => {
    if(a.length == 0) return b.length; 
    if(b.length == 0) return a.length; 

    var matrix = [];

    // increment along the first column of each row
    var i;
    for(i = 0; i <= b.length; i++){
        matrix[i] = [i];
    }

    // increment each column in the first row
    var j;
    for(j = 0; j <= a.length; j++){
        matrix[0][j] = j;
    }

    // Fill in the rest of the matrix
    for(i = 1; i <= b.length; i++){
        for(j = 1; j <= a.length; j++){
        if(b.charAt(i-1) == a.charAt(j-1)){
            matrix[i][j] = matrix[i-1][j-1];
        } else {
            matrix[i][j] = Math.min(matrix[i-1][j-1] + 1, // substitution
                                    Math.min(matrix[i][j-1] + 1, // insertion
                                            matrix[i-1][j] + 1)); // deletion
        }
        }
    }

    return matrix[b.length][a.length];
};

const matchTitles = async (s1,s2) => {

    const temp1 = s1.split("(");
    const temp2 = temp1[0].split(":");
    const temp3 = temp2[0].split("-");

    const diff  = await LevenshteinDistance(temp3[0].toLowerCase(),s2.toLowerCase());
    const percentSimilar = 1 - (diff/s2.length);

    return percentSimilar >= 0.9;
}


const getDetails = async (url) => {

    const browser = await puppeteer.launch({headless: false});
    const page = await browser.newPage();
    await page.goto(url, {timeout: 0});

    const bookData = await page.evaluate(async () => {
        var price = Array.from(document.querySelectorAll('.payBlkBig'),(e) => e.innerHTML)
        var author = Array.from(document.querySelectorAll('.p-keyfeatures ul>li:nth-child(5) .h-content'),(e) => e.innerHTML)
        var publisher = Array.from(document.querySelectorAll('.p-keyfeatures ul>li:nth-child(3) .h-content'),(e) => e.innerHTML)
        
        return { rate: price[0], auth: author[0], pub: publisher[0]};
    })
    await browser.close()
    return bookData;
}
const getTitlesUrl = async (url) => {
    
    const bookData = await page.evaluate(async () => {
        // No	Book Title	ISBN	Found	URL	Price	Author	Publisher	In Stock
        var snapTitles = Array.from(document.querySelectorAll('.product-title'), (e) => e.title)
        var snapUrl = Array.from(document.querySelectorAll('.product-desc-rating .noUdLine'),(e) => e.href)
        
        return {titles: snapTitles,urls: snapUrl};
    })
    await browser.close()
    return bookData;
}


const writeToExcel = (jsonDataArr) => {
    const newWB = XLSX.utils.book_new();
    const newWS = XLSX.utils.json_to_sheet(jsonDataArr)
    XLSX.utils.book_append_sheet(newWB,newWS,"Sheet1");
    XLSX.writeFile(newWB,"Output.xlsx");
}

const generate = () => {
    var dataArr = [];
    for (let index = 0; index < data.length; index++) {
        const isbn = data[index].ISBN;
        const title = data[index]['Book Title'];
        var url = `https://www.snapdeal.com/search?keyword=${isbn}&sort=plth`
        
        getTitlesUrl(url).then((json) => {
            const arr = json["titles"];
            let bookFoundAt = -1;
            var bookTitle;
            for(let i = 0; i < arr.length; i++){
                var t1 = arr[i];
                var res = matchTitles(t1,title)
                if(res){
                    bookTitle = t1;
                    bookFoundAt = i;
                    break;
                } 
            }
            if(bookFoundAt != -1){
                const pageUrl = json["urls"].at(bookFoundAt)
                let jsonData;
                getDetails(pageUrl).then(jso => {
                    var givenprice = jso["rate"]
                    var ath = jso["auth"]
                    var publisher = jso["pub"]
                    jsonData = {BookTitle: bookTitle, ISBN:	isbn, Found: 'yes' ,URL: pageUrl, Price: givenprice,Author: ath ,Publisher: publisher,	InStock: 'yes'}
                    dataArr[dataArr.length] = jsonData
                })
            }  
            else console.log("not found")
        })
    }
    return (dataArr);
    // writeToExcel(dataArr)
    
}

generate((res) => {
    console.log(res);
});



