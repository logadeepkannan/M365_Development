const express = require('express');
const fetch = require('node-fetch');
const app = express();
const port = 3000;

app.use(express.static('public')); // Serve static files from the 'public' directory

app.get('/projects', async (_req, res) => {
    const requestUrl = `https://ittridentsqa.sharepoint.com/sites/TridentSQAM365InternalSolution/_api/web/lists/getbytitle('Project_Dropdown')/items?$select=Title`;

    try {
        const response = await fetch(requestUrl, {
            method: 'GET',
            headers: {
                "Access-Control-Allow-Origin": "*",
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json',
                "odata.context": "https://graph.microsoft.com/v1.0/$metadata#sites('8d080d29-ef53-4995-87cd-3b7ab212e3f0%2C9b16a920-40c0-4e26-9b65-b5337c6896bc')/lists('%7Ba93b31bd-1c0f-4c9d-84e3-a36b3cabcfa6%7D')/items(fields(Title))",
                'odata-version': ''
            }
        });

        if (response.ok) {
            const data = await response.json();
            res.json(data.d.results);
        } else {
            res.status(response.status).send('Error fetching project data');
        }
    } catch (error) {
        res.status(500).send('Error: ' + error.message);
    }
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});
