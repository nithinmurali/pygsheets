Authorizing pygsheets
=====================

There are multiple ways to authorize google sheets. First you should create a developer account (follow below steps) and
create the type of credentials depending on your need. But remember not to give away any of there credentials, as your
usage quota is limited.


1. Head to `Google Developers Console <https://console.developers.google.com>`_ and create a new project (or select the one you have.)

.. image:: https://raw.githubusercontent.com/nithinmurali/tmpdatas/master/pygsheets/images/new_proj.png
    :alt: NEW PROJ

2.  You will be redirected to the API Manager, there Under "Library", Google APIs click on "Sheets API".

.. image:: https://raw.githubusercontent.com/nithinmurali/tmpdatas/master/pygsheets/images/apis.png
    :alt: APIs


3. In the API screen click on 'ENABLE' to enable this API

.. image:: https://raw.githubusercontent.com/nithinmurali/tmpdatas/master/pygsheets/images/api_enable.png
    :alt: Enabled APIs


4. Similarly enable the "Drive API". We require drives api for getting list of spreadsheets, deleting them etc.

Now you have to choose the type of credential you want to use. For this you have following two options:

Signed Credentials
------------------
In this option you will be given an unique email, and your application will be able to acesss all the sheets shared with that
email. No Authentication will be required in this case.


5. Go to "Credentials" Tab and choose "Create Credentials > Service Account Key".

6. Now choose the service account as 'App Engine default' and Key type as JSON and click create

.. image:: https://raw.githubusercontent.com/nithinmurali/tmpdatas/master/pygsheets/images/new_service_key.png
    :alt: Google Developers Console

You will automatically download a JSON file with this data.

.. image:: https://raw.githubusercontent.com/nithinmurali/tmpdatas/master/pygsheets/images/service_key_created.png
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


5. First you need to configure how the conscent sceen will look while asking for autorization. Go to "Credentials" Side Tab and choose "OAuth Conscent screen". There inset all the data you need to show while asking for authorization and save it

.. image:: https://raw.githubusercontent.com/nithinmurali/tmpdatas/master/pygsheets/images/oauth_conscent.png
    :alt: OAuth Conscent


6. Go to "Credentials" Tab and choose "Create Credentials > OAuth Client ID".

.. image:: https://raw.githubusercontent.com/nithinmurali/tmpdatas/master/pygsheets/images/creds_choose.png
    :alt: Google Developers Console

7. Now choose the Application Type as 'Other'

.. image:: https://raw.githubusercontent.com/nithinmurali/tmpdatas/master/pygsheets/images/create_client.png
    :alt: Client ID type


Now click on the download button to download the 'client_secretxxx.json' file

.. image:: https://raw.githubusercontent.com/nithinmurali/tmpdatas/master/pygsheets/images/download_client.png
    :alt: Download Credentials JSON from Developers Console


8. Find the client_id from the file, your application will be able to acess any sheet which is shared with this email. To use this file initiliaze the pygsheets client as shown
::

    gc = pygsheets.authorize(outh_file='client_secretxxx.json')

::

First time this will ask you to authorize pygsheets to acess your google sheets and drive, for this it will open a tab
in the brower, where you have to provide your google credentials and authorize it. This will create a json file with the
tokens based on the `outh_creds_store` param. So that you dont have to authorize it everytime you run the application.
In case if you already have a file with tokens then you can just pass it as the outh_file instead of the client secret file.

Incase you are running the script in a headless brower where it can't open a broweser, you can enbale non-local authorization.
Hence instead of opening a brower in the same meachine, it will provide a link which you can run on your local computer
and authorize the application.

::

    gc = pygsheets.authorize(outh_file='client_secretxxx.json', outh_nonlocal=True)


Custom Credentials Objects
--------------------------
If you have another method of authenicating you can easily create a custom credentials object.

::

    class Credentials (object):
        def __init__ (self, access_token=None):
            self.access_token = access_token

        def refresh (self, http):
            # get new access_token
            # this only gets called if access_token is None

Then you could pass this for authorization as

::

    gc = pygsheets.authorize(credentials=mycreds)

