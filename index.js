const cors = require('cors');
const { MongoClient } = require('mongodb');
const objectId = require('mongodb').ObjectId
require('dotenv').config();
const port = process.env.PORT || 5000
var xlsx = require('node-xlsx');
const multer = require('multer')
const path = require('path');
const app = require('express')();
const express = require('express');
const excelToJson = require('convert-excel-to-json');
const fs = require("fs")
const uri = `mongodb+srv://${process.env.DB_USER}:${process.env.DB_PASS}@cluster0.wanl6.mongodb.net/myFirstDatabase?retryWrites=true&w=majority`;
const client = new MongoClient(uri, { useNewUrlParser: true, useUnifiedTopology: true });


var bodyParser = require('body-parser');
// MIDDLEWARE
app.use(express.json())
app.use(cors());
app.use(express.static("uploads"));
app.use(bodyParser.json({ limit: "50mb" }));
app.use(bodyParser.urlencoded({ limit: "50mb", extended: true, parameterLimit: 50000 }));

const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, '');
    },

    filename: function (req, file, cb) {
        cb(null, file.originalname.split(".")[0] + '-' + Date.now() + path.extname(file.originalname));
    }
});

const upload = multer({ storage: storage })







async function run() {
    try {
        await client.connect();
        const database2 = client.db("unityMart");
        const unityMartProductsCollection = database2.collection("products");


        app.post("/upload-excel", upload.single('file'), async (req, res) => {
            const vendor = req.body
            const file = req.file;
            const filename = file.filename;
            // const currPath = fs.readdirSync("./uploads/" + filename)
            // // const file = "./uploads/" + filename
            // console.log(currPath);

            if ('/uploads/' + filename) {
                // const currPath = fs.readdirSync("./uploads/" + filename)
                // const file = "./uploads/" + filename
                // console.log(currPath);
                const worksheetsArray = xlsx.parse(filename);
                const excelData = excelToJson({
                    sourceFile: filename,
                    sheets: [{
                        // Excel Sheet Name
                        name: worksheetsArray[0].name,

                        // Header Row -> be skipped and will not be present at our result object.
                        header: {
                            rows: 1
                        },

                        // Mapping columns to keys
                        columnToKey: {
                            B: 'title',
                            C: 'brand',
                            D: 'reg_price',
                            E: 'sale_price',
                            F: 'stock',
                            G: 'images',
                            H: 'categories',
                            I: 'product_des',
                            J: 'categories',
                        }
                    }]
                });
                // const result = excelData.Customers.map(row => row)
                const withImages = excelData.Customers.map((customer) => {
                    // const imgIds = customer.image_id.split(",");
                    const imgSrcs = customer.images?.split(",");
                    const categories = customer.categories?.split(",");
                    return {
                        title: customer.title || null,
                        brand: customer.brand || null,
                        reg_price: customer.reg_price.toString() || null,
                        sale_price: customer.sale_price.toString() || null,
                        stock: customer.stock.toString() || null,
                        categories: categories.map((id, i) => {
                            return { label: categories[i], value: categories[i].split(" ").join("-") }
                        }),
                        product_des: customer.product_des || null,
                        images: imgSrcs.map((id, i) => {
                            return { id: i + 1, src: imgSrcs[i] };
                        }),
                        offerDate: "",
                        attributes: [],
                        publisherDetails: {}
                    };
                });
                const result = await unityMartProductsCollection.insertMany(withImages)
                if (result.insertedCount) {
                    res.send(withImages)
                }
                console.log(withImages);
            }
            // const worksheetsArray = xlsx.parse('/uploads/' + filename); // parses a file


        });

    } finally {
        //   await client.close();
    }
}
run().catch(console.dir);






app.get('/', (req, res) => {
    res.send('Hello World!')
})

app.listen(port, () => {
    console.log(`Example app listening at http://localhost:${port}`)
})