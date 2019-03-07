Authorizing pygsheets
=====================

There are multiple ways to authorize google sheets. First you should create a developer account (follow below steps) and
create the type of credentials depending on your need. These credentials give the python script access to a google account.
 But remember not to give away any of these credentials, as your usage quota is limited.


1. Head to `Google Developers Console <https://console.developers.google.com>`_ and create a new project (or select the one you have.)

.. image:: https://raw.githubusercontent.com/nithinmurali/tmpdatas/master/pygsheets/images/new_proj1.png
    :alt: NEW PROJ

.. image:: https://raw.githubusercontent.com/nithinmurali/tmpdatas/master/pygsheets/images/new_proj2.png
    :alt: NEW PROJ ADD


2.  You will be redirected to the Project Dashboard, there click on "Enable Apis and services", search for "Sheets API".

.. image:: https://raw.githubusercontent.com/nithinmurali/tmpdatas/master/pygsheets/images/apis.png
    :alt: APIs


3. In the API screen click on 'ENABLE' to enable this API

.. image:: https://raw.githubusercontent.com/nithinmurali/tmpdatas/master/pygsheets/images/api_enable.png
    :alt: Enabled APIs


4. Similarly enable the "Drive API". We require drives api for getting list of spreadsheets, deleting them etc.

Now you have to choose the type of credential you want to use. For this you have following two options:

OAuth Credentials
-----------------
This is the best option if you are trying to edit the spreadsheet on behalf of others. Using this method
you the script can get access to all the spreadsheets of the account. The authorization process (giving
password and email) has to be completed only once. Which will grant the application full access to all of the
users sheets. Follow this procedure below to get the client secret:

 .. note::
        Make sure to not share the created authentication file with anyone, as it will give direct access
        to your enabled APIs.


5. First you need to configure how the consent screen will look while asking for authorization.
Go to "Credentials" side tab and choose "OAuth Consent screen". Input all the required data in the form.

.. image:: https://raw.githubusercontent.com/nithinmurali/tmpdatas/master/pygsheets/images/oauth_conscent.png
    :alt: OAuth Consent


6. Go to "Credentials" tab and choose "Create Credentials > OAuth Client ID".

.. image:: https://raw.githubusercontent.com/nithinmurali/tmpdatas/master/pygsheets/images/creds_choose.png
    :alt: Google Developers Console

7. Next choose the application type as 'Other'

.. image:: https://raw.githubusercontent.com/nithinmurali/tmpdatas/master/pygsheets/images/create_client.png
    :alt: Client ID type


8. Next click on the download button to download the 'client_secret[...].json' file, make sure to remember where
you saved it:

.. image:: https://raw.githubusercontent.com/nithinmurali/tmpdatas/master/pygsheets/images/download_client.png
    :alt: Download Credentials JSON from Developers Console


9. By default authorize() expects a filed named 'client_secret.json' in the current working directory. If you did not
save the file there and renamed it, make sure to set the path:
::

    gc = pygsheets.authorize(client_secret='path/to/client_secret[...].json')

The first time this will ask you to complete the authentication flow. Follow the instructions in the console to
complete. Once completed a file with the authentication token will be stored in your current working
directory (to change this set credentials_directory). This file is used so that you don't have to authorize it
every time you run the application. So if you need to authorize script again you don't need the
client_secret but just this generated json file will do (pass its path as credentials_directory).

Please note that credentials_directory will override your client_secrect. So if you keep getting logged in
with some other account even when you are passing your accounts client_secrect, credentials_directory might be
the culprit.


Service Account
---------------
A service account is an account associated with an email. The account is authenticated with a pair of
public/private key making it more secure than other options. For as long as the private key stays private.
To create a service account follow these steps:

5. Go to "Credentials" tab and choose "Create Credentials > Service Account Key".

6. Next choose the service account as 'App Engine default' and Key type as JSON and click create:

.. image:: https://raw.githubusercontent.com/nithinmurali/tmpdatas/master/pygsheets/images/new_service_key.png
    :alt: Google Developers Console

7. You will now be prompted to download a .json file. This file contains the necessary private key for
account authorization. Remember where you stored the file and how you named it.

.. image:: https://raw.githubusercontent.com/nithinmurali/tmpdatas/master/pygsheets/images/service_key_created.png
    :alt: Download Credentials JSON from Developers Console

This is how this file may look like::

    {
        "type": "service_account",
        "project_id": "p....sdf",
        "private_key_id": "48.....",
        "private_key": "-----BEGIN PRIVATE KEY-----\nNrDyLw … jINQh/9\n-----END PRIVATE KEY-----\n",
        "client_email": "p.....@appspot.gserviceaccount.com",
        "client_id": "10.....454",
    }

7. The authorization process can be completed without any further interactions::

    gc = pygsheets.authorize(service_file='path/to/service_account_credentials.json')

Custom Credentials Objects
--------------------------
You can create your own authentication method and pass them like this::

    gc = pygsheets.authorize(custom_credentials=my_credentials)

This option will ignore any other parameters.
