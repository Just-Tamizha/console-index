const axios = require('axios').default;
const https = require('https');
const { google } = require('googleapis');
const fs = require('fs'); const path = require('path');
const xlsx = require('xlsx');
const getUserAccessDetails = JSON.parse(fs.readFileSync(path.join(__dirname, 'key.json')));

// Constants
const SCOPES = ['https://www.googleapis.com/auth/indexing'];
const ENDPOINT = 'https://indexing.googleapis.com/v3/urlNotifications:publish';
async function getAccessToken(client_email, private_key) {
    try {
        const auth = new google.auth.GoogleAuth({
            credentials: {
                client_email: client_email,
                private_key: private_key,
            }, scopes: SCOPES
        });
        const accessToken = await auth.getAccessToken();
        return { status: true, data: { token: accessToken } }
    } catch (error) {
        return { status: false, message: "Service account access Denied !" }
    }
}

async function googleIndex(http, url) {
    const content = { url: url.trim(), type: 'URL_DELETED' };
    try {
        const response = await axios.post(ENDPOINT, content, {
            headers: { Authorization: `Bearer ${http}` },
            httpsAgent: new https.Agent({ rejectUnauthorized: false }),
        });
        return { status: true, data: response.data }
    } catch (error) {
        if (error.response && error.response.status === 429) {
            return { status: false, message: 'Rate limit exceeded', type: "LIMIT" };
        } else {
            return { status: false, message: 'Server Disconnected after multiple retries!', type:"ERROR"};
        }
    }
}
const indexUrl = async () => {
    try {
        const http = await getAccessToken(getUserAccessDetails.client_email, getUserAccessDetails.private_key);
        if (!http.status) return console.log(`ERROR: ${http.message}`);
        const excelFilePath = path.join(__dirname, 'index.xlsx'); const workbook = xlsx.readFile(excelFilePath);
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
        let ok_url = 0; let limit_url = 0;let fail_url=0;
        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            if (limit_url > 2) { console.log(`\nLimit Reached !`); break; }
            if (row[1] == "" || row[1] == undefined || row[1] == null) {
                const result = await googleIndex(http.data.token, row[0]);
                console.log(`Index [${i}]: ${result.status}`)
                if (result.status) {row[1] = "OK"; ok_url++;}
                if (!result.status) {
                    row[2] = result.message; 
                    if(result.type =="LIMIT") limit_url++;
                    if(result.type =="ERROR") fail_url++;
                }
            }
        }
        const updatedSheet = xlsx.utils.aoa_to_sheet(data);
        workbook.Sheets[sheetName] = updatedSheet;
        xlsx.writeFile(workbook, excelFilePath);
        console.log(`\n------------------------------ \n \nTotal URL: ${data.length} \nSUCCESS: ${ok_url} \nFAILURE: ${data.length - ok_url}\n\n------------------------------`);
        console.log("COMPLETED !\n\n");

    } catch (error) {
        console.log(error)
    }

}
indexUrl()