//Modified implementation posted by @CrayonConstantinople here: https://www.reddit.com/r/spreadsheets/comments/5zbnbb/help_populate_cells_with_book_details_from_isbn/dfcz1kj/?utm_source=share&utm_medium=web3x&utm_name=web3xcss&utm_term=1&utm_content=share_button

s = SpreadsheetApp.getActiveSheet();

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Book Menu')
        .addItem('Get Book Details From ISBN', 'getBookDetailsFromISBN')
        .addItem('Get Book Details From Title', 'getBookDetailsFromTitle')
        .addToUi();
}

function getBookDetailsFromISBN(isbn) {
    // Query the book database by ISBN code.
    activeCell = s.getActiveCell();
    value = activeCell.getValue();
    isbn = isbn || value.toString(); // Steve Jobs book 

    // Not a valid ISBN if not 13 or 10 digits long.
    if (isbn.match(/(\d{13}|\d{10})/) == null) {
        throw new Error("Not a valid ISBN: " + isbn);
    }

    //URL fix, needed to add in region: https://stackoverflow.com/questions/11232691/google-books-api-server-not-accepting-calls-from-heroku-server
    var url = "https://www.googleapis.com/books/v1/volumes?country=US&q=isbn:" + isbn;
    var response = UrlFetchApp.fetch(url);
    var results = JSON.parse(response);
    if (results.totalItems) {

        // There'll be only 1 book per ISBN
        var book = results.items[0];

        var title = (book["volumeInfo"]["title"]);
        var subtitle = (book["volumeInfo"]["subtitle"]) || "*No Subtitle";
        var authors = (book["volumeInfo"]["authors"]);
        var printType = (book["volumeInfo"]["printType"]);
        var pageCount = (book["volumeInfo"]["pageCount"]);
        var publisher = (book["volumeInfo"]["publisher"]);
        var publishedDate = (book["volumeInfo"]["publishedDate"]);
        var webReaderLink = (book["accessInfo"]["webReaderLink"]);

        //Logger.log(book);
        results = [[title, subtitle, authors, printType, pageCount, publisher, publishedDate, webReaderLink]];

    } else {
        results = [["-", "-", "-", "-", "-", "-", "-", "-"]];
    }
    s.getRange(activeCell.getRow(), activeCell.getColumn() + 1, 1, results[0].length).setValues(results);
}

function getBookDetailsFromTitle(title) {
    // Query the book database by ISBN code.
    activeCell = s.getActiveCell();
    value = activeCell.getValue();
    title = title || value.toString();
    title = title.replaceAll(' ', '%20');

    //URL fix, needed to add in region: https://stackoverflow.com/questions/11232691/google-books-api-server-not-accepting-calls-from-heroku-server
    var url = "https://www.googleapis.com/books/v1/volumes?country=US&q=title:" + title;
    var response = UrlFetchApp.fetch(url);
    var results = JSON.parse(response);
    if (results.totalItems) {

        // There'll be only 1 book per ISBN
        var book = results.items[0];

        var isbn = (book["volumeInfo"]["industryIdentifiers"][1]["identifier"])
        var subtitle = (book["volumeInfo"]["subtitle"]) || "*No Subtitle";
        var authors = (book["volumeInfo"]["authors"]);
        var printType = (book["volumeInfo"]["printType"]);
        var pageCount = (book["volumeInfo"]["pageCount"]);
        var publisher = (book["volumeInfo"]["publisher"]);
        var publishedDate = (book["volumeInfo"]["publishedDate"]);
        var webReaderLink = (book["accessInfo"]["webReaderLink"]);

        //Logger.log(book);
        results = [[subtitle, authors, printType, pageCount, publisher, publishedDate, webReaderLink]];

    } else {
        results = [["-", "-", "-", "-", "-", "-", "-", "-"]];
    }

    s.getRange(activeCell.getRow(), activeCell.getColumn() - 1, 1, 1).setValue(isbn)
    s.getRange(activeCell.getRow(), activeCell.getColumn() + 1, 1, results[0].length).setValues(results);
}