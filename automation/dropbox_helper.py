import dropbox

app_key = '8w7ts322hfkb9ek'
app_secret = '8dp228du8qsvj46'
access_token = 'kSVIwQWSbHMAAAAAAAAL64PmaxHp3J-LHwFp-f0XC9J2nx5Ef_MCNHYGbFAeG2LA'


def upload_file_to_dropbox(file_location, filename):
    metadata = None
    f = open('working-draft.txt', 'rb')
    dbx = dropbox.Dropbox(access_token)
    dbx.users_get_current_account()
    metadata = dbx.files_upload(f, str(filename))

    return metadata
