const axios = require('axios');
const { parseString } = require('xml2js');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const xlsx = require('xlsx');
require('dotenv').config();

const urls = [
    'https://9f4a-2409-40f2-104b-6ea6-c177-bb72-871d-9019.ngrok-free.app',
    'https://bbf0-2406-7400-111-9753-2cd0-9b93-f25d-178c.ngrok-free.app'
];

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
                            PARTYLEDGERNAME: v.PARTYLEDGERNAME ? v.PARTYLEDGERNAME[0] : '',
                            VOUCHERNUMBER: v.VOUCHERNUMBER ? v.VOUCHERNUMBER[0] : '',
                            AMOUNT: v["ALLLEDGERENTRIES.LIST"][0]?.AMOUNT?.[0] || 0,
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

// Export handler for Vercel
module.exports = async (req, res) => {
  if (req.method === 'POST') {
      // Existing POST logic
      const excelBuffer = await fetchDataAndParseToExcel();
      await sendEmail(excelBuffer);
      res.status(200).send('Report generated and email sent successfully.');
  } else if (req.method === 'GET') {
      res.status(200).send('Server is running. Use POST for data processing.');
  } else {
      res.status(405).send('Method Not Allowed');
  }
};

