const express = require('express');
const axios = require('axios');
const { parseString } = require('xml2js');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const xlsx = require('xlsx');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000; // Render provides the PORT

const urls = [
    ' https://54a6-2406-7400-111-9753-c67-163b-453a-c034.ngrok-free.app',
    'https://f5c9-14-195-101-178.ngrok-free.app'
];
const payload = `
<ENVELOPE>
    <HEADER>
        <VERSION>1</VERSION>
        <TALLYREQUEST>Export</TALLYREQUEST>
        <TYPE>Data</TYPE>
        <ID>Day Book</ID>
    </HEADER>
    <BODY>
        <DESC>
            <STATICVARIABLES>
                <EXPLODEFLAG>Yes</EXPLODEFLAG>
                <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
            </STATICVARIABLES>
            <TDL>
                <TDLMESSAGE>
                    <REPORT NAME="Day Book" ISMODIFY="Yes">
                        <ADD>Set : SV From Date:"20240401"</ADD>
                        <ADD>Set : SV To Date:"20240405"</ADD>
                        <ADD>Set : ExplodeFlag : Yes</ADD>
                    </REPORT>
                </TDLMESSAGE>
            </TDL>
        </DESC>
    </BODY>
</ENVELOPE>`;

const config = {
    auth: {
        clientId: process.env.AZURE_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
        clientSecret: process.env.AZURE_SECRET,
    },
};

const cca = new ConfidentialClientApplication(config);

const getAccessToken = async () => {
    try {
        const result = await cca.acquireTokenByClientCredential({
            scopes: [process.env.AZURE_SCOPES],
        });
        return result.accessToken;
    } catch (error) {
        console.error('Error fetching access token:', error.message);
        throw error;
    }
};

const aggregateData = (vouchers) => {
    const aggregatedVouchers = {};

    vouchers.forEach(voucher => {
        const key = `${voucher.DATE}-${voucher.VOUCHERNUMBER}-${voucher.PARTYLEDGERNAME}`;
        if (!aggregatedVouchers[key]) {
            aggregatedVouchers[key] = { ...voucher, AMOUNT: 0 };
        }
        aggregatedVouchers[key].AMOUNT += parseFloat(voucher.AMOUNT || 0);
    });

    return Object.values(aggregatedVouchers);
};

const fetchDataAndParseToExcel = async () => {
    const allVouchers = [];

    for (const url of urls) {
        try {
            const response = await axios.post(url, payload, {
                headers: { 'Content-Type': 'text/xml' },
            });

            parseString(response.data, (err, result) => {
                if (err) {
                    console.error('Error parsing XML:', err);
                    return;
                }

                const tallyMessages = result.ENVELOPE?.BODY[0]?.DATA[0]?.TALLYMESSAGE || [];
                tallyMessages.forEach(tallyMessage => {
                    if (tallyMessage.VOUCHER) {
                        const v = tallyMessage.VOUCHER[0];
                        allVouchers.push({
                            DATE: v.DATE ? v.DATE[0] : '',
                            GUID: v.GUID ? v.GUID[0] : '',
                            NARRATION: v.NARRATION ? v.NARRATION[0] : '',
                            OBJECTUPDATEACTION: v.OBJECTUPDATEACTION ? v.OBJECTUPDATEACTION[0] : '',
                            GSTREGISTRATION: v.GSTREGISTRATION ? v.GSTREGISTRATION[0]._ : '',
                            VOUCHERTYPENAME: v.VOUCHERTYPENAME ? v.VOUCHERTYPENAME[0] : '',
                            PARTYLEDGERNAME: v.PARTYLEDGERNAME ? v.PARTYLEDGERNAME[0] : '',
                            VOUCHERNUMBER: v.VOUCHERNUMBER ? v.VOUCHERNUMBER[0] : '',
                            CMPGSTREGISTRATIONTYPE: v.CMPGSTREGISTRATIONTYPE ? v.CMPGSTREGISTRATIONTYPE[0] : '',
                            CMPGSTSTATE: v.CMPGSTSTATE ? v.CMPGSTSTATE[0] : '',
                            NUMBERINGSTYLE: v.NUMBERINGSTYLE ? v.NUMBERINGSTYLE[0] : '',
                            CSTFORMISSUETYPE: v.CSTFORMISSUETYPE ? v.CSTFORMISSUETYPE[0] : '',
                            FBTPAYMENTTYPE: v.FBTPAYMENTTYPE ? v.FBTPAYMENTTYPE[0] : '',
                            PERSISTEDVIEW: v.PERSISTEDVIEW ? v.PERSISTEDVIEW[0] : '',
                            VOUCHERKEY: v.VOUCHERKEY ? v.VOUCHERKEY[0] : '',
                            LEDGERNAME: v["ALLLEDGERENTRIES.LIST"][0]?.LEDGERNAME?.[0] || '',
                            AMOUNT: v["ALLLEDGERENTRIES.LIST"][0]?.AMOUNT?.[0] || ''
                        });
                    }
                });
            });
        } catch (error) {
            console.error(`Error fetching data from ${url}:`, error.message);
        }
    }

    if (allVouchers.length === 0) {
        console.error('No valid VOUCHER entries found.');
        return;
    }

    const aggregatedVouchers = aggregateData(allVouchers);
    const wb = xlsx.utils.book_new();
    const ws = xlsx.utils.json_to_sheet(aggregatedVouchers);
    xlsx.utils.book_append_sheet(wb, ws, 'Aggregated Data');

    return xlsx.write(wb, { bookType: 'xlsx', type: 'buffer' });
};

const sendEmail = async (excelBuffer) => {
    const accessToken = await getAccessToken();

    const emailData = {
        message: {
            subject: "Test Email with Excel Attachment",
            body: {
                contentType: "Text",
                content: "Hello, this is a test email sent using Microsoft Graph API with an Excel attachment!",
            },
            toRecipients: [
                { emailAddress: { address: "satyamankit13@gmail.com" } },
            ],
            attachments: [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    name: "output.xlsx",
                    contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    contentBytes: excelBuffer.toString('base64'),
                },
            ],
        },
    };

    await axios.post(
        `https://graph.microsoft.com/v1.0/users/${process.env.AZURE_USER_ID}/sendMail`,
        emailData,
        {
            headers: {
                Authorization: `Bearer ${accessToken}`,
                'Content-Type': 'application/json',
            },
        }
    );
};

// Define Express routes
app.post('/process', async (req, res) => {
    try {
        const excelBuffer = await fetchDataAndParseToExcel();
        await sendEmail(excelBuffer);
        res.status(200).send('Report generated and email sent successfully.');
    } catch (error) {
        console.error('Error during process execution:', error.message);
        res.status(500).send('An error occurred during the process.');
    }
});

app.get('/', (req, res) => {
    res.send('Server is running. Use POST /process for data processing.');
});

// Start Express server
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
