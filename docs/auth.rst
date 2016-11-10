Authorizing pygsheets
=====================

There are multiple ways to authorize google sheets. For all type of credentials you should create a developer account (follow below steps)
but remember not to give away any of there credentials, as your usage quota is limited.

::

1. Head to `Google Developers Console <https://console.developers.google.com/project>`_ and create a new project (or select the one you have.)

2.  You will be redirected to the API Manager, there Under "Library", Google APIs click on "Sheets API".

.. image:: https://cloud.githubusercontent.com/assets/264674/7033107/72b75938-dd80-11e4-9a9f-54fb10820976.png
    :alt: Enabled APIs

3. In the API screen click on 'ENABLE' to enable this API

.. image:: https://cloud.githubusercontent.com/assets/264674/7033107/72b75938-dd80-11e4-9a9f-54fb10820976.png
    :alt: Enabled APIs

4. Similerly enable the  and "Drive API". We require drives api for getting list of spreadsheets.


Signed Credentials
-----------------
In this option you will be given an unuque email, and your sheet will be able to acesss all the sheets shared with that
email. No Authentication will be required in this case.

::

5. Go to "Credentials" Tab and choose "Create Credentials > Service Account Key".

.. image:: https://cloud.githubusercontent.com/assets/1297699/12098271/1616f908-b319-11e5-92d8-767e8e5ec757.png
    :alt: Google Developers Console

You will automatically download a JSON file with this data.

.. image:: https://cloud.githubusercontent.com/assets/264674/7033081/3810ddae-dd80-11e4-8945-34b4ba12f9fa.png
    :alt: Download Credentials JSON from Developers Console

This is how this file may look like:

::

    {
        "type": "service_account",
        "project_id": "p....sdf",
        "private_key_id": "48.....",
        "private_key": "-----BEGIN PRIVATE KEY-----\nNrDyLw … jINQh/9\n-----END PRIVATE KEY-----\n",
        "client_email": "p.....@appspot.gserviceaccount.com",
        "client_id": "10.....454",
    }



6. Find the client_id from the file, your application will be able to acess any sheet which is shared with this email. To use this file initiliaze the pygsheets client as shown
::

    gc = pygsheets.authorize(service_file='service_creds.json')


OAuth Credentials
-----------------
This is the best option if you are trying to edit the spreadsheet on behalf of others. Here for the first time the user will
be asked to authenticate your application. From therafter the application can acess all his spreadsheets. For using this
you will need 'OAuth client ID' file. Follow this procedure below to generate it -

::


5. First you need to configure how the conscent sceen will look while asking for autorization. Go to "Credentials" Side Tab and choose "OAuth Conscent screen". There inset all the data you need to show while asking for authorization and save it

.. image:: https://cloud.githubusercontent.com/assets/1297699/12098271/1616f908-b319-11e5-92d8-767e8e5ec757.png
    :alt: OAuth Conscent


6. Go to "Credentials" Tab and choose "Create Credentials > OAuth Client ID".

.. image:: https://cloud.githubusercontent.com/assets/1297699/12098271/1616f908-b319-11e5-92d8-767e8e5ec757.png
    :alt: Google Developers Console

7. Now choose the Application Type as 'Other'

.. image:: https://cloud.githubusercontent.com/assets/1297699/12098271/1616f908-b319-11e5-92d8-767e8e5ec757.png
    :alt: Client ID type


Now click on the download button to download the 'client_secretxxx.json' file

.. image:: https://cloud.githubusercontent.com/assets/264674/7033081/3810ddae-dd80-11e4-8945-34b4ba12f9fa.png
    :alt: Download Credentials JSON from Developers Console


8. Find the client_id from the file, your application will be able to acess any sheet which is shared with this email. To use this file initiliaze the pygsheets client as shown
::

    gc = pygsheets.authorize(oauth_file='client_secretxxx.json')

