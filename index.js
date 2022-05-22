const express = require('express');
const { default:axios} = require('axios');
const fileUpload = require('express-fileupload');
const reader = require('xlsx');
var Readable = require('stream').Readable;


const app = express();

app.use(express.static("public"));
app.use(fileUpload());

// converts the buffer xlsx file to stream
const bufferToStream = buffer => { 
  var stream = new Readable();
  stream.push(buffer);
  stream.push(null);

  return stream;
}

// get price for the product with the product_id
const getPrice = async product_id => {
  if(product_id && product_id.length === 0) return NaN;
  try {
    const res = await axios.get(`https://api.storerestapi.com/products/${product_id}`);
    const { data } = res.data;
    return data['price'];
  } catch (err) {
    // This would be added to instead of price so that incase there is a product for which 
    // the above api gives an error it would be noified in the excel
    return 'An error occurred with the product api';
  }
  
}


// upload api which adds price to the excel and then returns to the client
app.post('/upload', async function(req, res) {
  if (!req.files || Object.keys(req.files).length === 0) {
    return res.status(400).send('No files were uploaded.');
  }

  try {
    // Uploaded file
    const excel = req.files.excel;
    console.log('Received file');
  // Extracting data from the file
  const file = reader.read(excel.data, ArrayBuffer);
  
    // Converting data to json
    const data = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[0]]);

    // Adding the price
    await Promise.all(data.map(async product => {
      const price = await getPrice(product['product_code']);
      product['price'] = price;
      return product;
    }));
    console.log('Data added successfully!');

    // converting data to sheet
    const sheet = reader.utils.json_to_sheet(data);
    // new workbook
    const workbook = reader.utils.book_new()
    // add sheet to a new workbook
    // here previous workbook (file) can be used and a sheet with price can be added
    reader.utils.book_append_sheet(workbook, sheet);

    const wbOpts = { bookType: "xlsx", type: "buffer" };

    // write workbook buffer
    const buffer = reader.write(workbook, wbOpts);                  
    console.log('Converted file to buffer!');

     // convert buffer to stream
    const stream = bufferToStream(buffer);                  
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader("Content-Disposition", "attachment; filename=products.xlsx");
    stream.pipe(res);  
    console.log('Process finished!');

  } catch (error) {
    res.status(500).send('An error occured at the server!!!');
  }
});


app.listen(3000, function() {
     console.log('Listening on port 3000');
});