//Modules
import express from 'express';
import bunyan from 'bunyan';
import bodyParser from 'body-parser';
import fetch from 'node-fetch';

//Load values from .env file
import dotenv from 'dotenv';
dotenv.config();

const app = express();
const log = bunyan.createLogger({ name: 'Resource Owner Password Credentials Flow' });

app.use(express.static('public'));

// parse application/x-www-form-urlencoded
app.use(bodyParser.urlencoded({ extended: false }));

app.set('view engine', 'ejs');

app.get('/', (req, res) => {
    res.render('index');
});

//Step 1: Get the access token
app.get('/get/the/token', (req, res) => {

    const Token_Endpoint = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`;
    const Grant_Type = 'password';    
    const Client_Id = process.env.CLIENT_ID;
    const Client_Secret = process.env.CLIENT_SECRET;
    const UserName = process.env.USER_NAME;
    const Password = process.env.USER_PASSWORD;
    const Scope = 'https://graph.microsoft.com/User.Read';

    let body = `grant_type=${Grant_Type}&client_id=${Client_Id}&client_secret=${Client_Secret}&username=${UserName}&password=${Password}&scope=${encodeURIComponent(Scope)}`;

    log.info(`Endpoint: ${Token_Endpoint}`);

    log.info(`Body: ${body}`);

    fetch(Token_Endpoint, {
        method: 'POST',
        body: body,
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'           
        }
    }).then(async response => {

        let json = await response.json();
        res.render('access-token', { token: JSON.stringify(json, undefined, 2) }); //you shouldn't share the access token with the client-side

    }).catch(error => {
        log.error(error.message);
    });
});

//Step 2: Call the protected API
app.post('/call/ms/graph', (req, res) => {

    let access_token = JSON.parse(req.body.token).access_token;

    const Microsoft_Graph_Endpoint = 'https://graph.microsoft.com/beta';
    const Acction_That_I_Have_Access_Because_Of_My_Scope = '/me';

    //Call Microsoft Graph with your access token
    fetch(`${Microsoft_Graph_Endpoint}${Acction_That_I_Have_Access_Because_Of_My_Scope}`, {
        headers: {
            'Authorization': `Bearer ${access_token}`
        }
    }).then(async response => {

        let json = await response.json();
        res.render('calling-ms-graph', { response: JSON.stringify(json, undefined, 2) });
    });
});

app.listen(8000);