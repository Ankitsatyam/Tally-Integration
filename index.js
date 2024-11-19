const axios = require('axios');
const { parseString } = require('xml2js');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const xlsx = require('xlsx');
require('dotenv').config();





// Define the payload to be sent in the request
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

// URLs to send the request to
const urls = [
    'https://9f4a-2409-40f2-104b-6ea6-c177-bb72-871d-9019.ngrok-free.app',
    'https://bbf0-2406-7400-111-9753-2cd0-9b93-f25d-178c.ngrok-free.app '
];

const config = {
  auth: {
      clientId: '4d606f55-2139-44b6-979b-ead2a257f591',
      authority: 'https://login.microsoftonline.com/617aa372-a362-423a-b00d-d33cecf52ea7',
      clientSecret: process.env.AZURE_SECRET,
  },
};

const cca = new ConfidentialClientApplication(config);

// Function to get the access token
const getAccessToken = async () => {
    try {
        const result = await cca.acquireTokenByClientCredential({
            scopes: ['https://graph.microsoft.com/.default'],
        });
        return result.accessToken;
    } catch (error) {
        console.error('Error fetching access token:', error.message);
        throw error;
    }
};

// Function to aggregate data based on your conditions
const aggregateData = (vouchers) => {
    const aggregatedVouchers = {};

    vouchers.forEach(voucher => {
        const key = `${voucher.DATE}-${voucher.VOUCHERNUMBER}-${voucher.PARTYLEDGERNAME}`;

        if (!aggregatedVouchers[key]) {
            aggregatedVouchers[key] = { ...voucher, AMOUNT: 0 };
        }

        aggregatedVouchers[key].AMOUNT += parseFloat(voucher.AMOUNT || 0);
    });

    return Object.values(aggregatedVouchers); // Return aggregated data
};

// Function to fetch and parse XML data
const fetchDataAndParseToExcel = async () => {
    const allVouchers = [];

    for (const url of urls) {
        try {
            const response = await axios.post(url, payload, {
                headers: { 'Content-Type': 'text/xml' }
            });

            parseString(response.data, (err, result) => {
                if (err) {
                    console.error('Error parsing the XML:', err);
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

    // Aggregate data based on your criteria
    const aggregatedVouchers = aggregateData(allVouchers);

    // Generate Excel file as a buffer (in-memory)
    const wb = xlsx.utils.book_new();
    const ws = xlsx.utils.json_to_sheet(aggregatedVouchers);
    xlsx.utils.book_append_sheet(wb, ws, 'Aggregated Data');

    // Generate the buffer without saving the file locally
    const excelBuffer = xlsx.write(wb, { bookType: 'xlsx', type: 'buffer' });

    return excelBuffer; // Return the buffer to be used in the email
};

// Function to send an email using Microsoft Graph API
const sendEmail = async (excelBuffer) => {
    try {
        const accessToken = await getAccessToken();

        const emailData = {
            message: {
                subject: "Test Email with Excel Attachment",
                body: {
                    contentType: "Text",
                    content: "Hello, this is a test email sent using Microsoft Graph API with an Excel attachment!",
                },
                toRecipients: [
                    {
                        emailAddress: {
                            address: "satyamankit13@gmail.com",
                        },
                    },
                ],
                attachments: [
                    {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        name: "output.xlsx", // Excel file name
                        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", // Excel MIME type
                        contentBytes: excelBuffer.toString('base64'), // Convert buffer to base64
                    },
                ],
            },
        };

        const userId = "ankit.k@forcecloudlabs.com";
        const response = await axios.post(
            `https://graph.microsoft.com/v1.0/users/${userId}/sendMail`,
            emailData,
            {
                headers: {
                    Authorization: `Bearer ${accessToken}`,
                    'Content-Type': 'application/json',
                },
            }
        );

        console.log('Email sent successfully:', response.status);
    } catch (error) {
        console.error('Error sending email:', error.response ? error.response.data : error.message);
    }
};

// Run the process (Excel generation and email sending)
(async () => {
    try {
        const excelBuffer = await fetchDataAndParseToExcel(); // Wait for Excel generation in buffer
        await sendEmail(excelBuffer); // Send email with in-memory Excel buffer
    } catch (error) {
        console.error('Error during process execution:', error.message);
    }
})();
