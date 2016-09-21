#!/usr/bin/env python3

import httplib2
import os
import logging

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

export_types = {
  'htm':  'text/html',
  'html': 'text/html',
  'txt':  'text/plain',
  'text': 'text/plain',
  'rtf':  'application/rtf',
  'pdf':  'application/pdf',
  'odf':  'application/vnd.oasis.opendocument.text',
  'doc':  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
}

logger = logging.getLogger()

# Turn down google api client logging, since INFO level is a bit too chatty
glogger = logging.getLogger('googleapiclient.discovery')
glogger.setLevel(logging.WARNING)

SCOPES='https://www.googleapis.com/auth/drive.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'drive-exporter'

try:
  import argparse
  parser = argparse.ArgumentParser(parents=[tools.argparser], description='Export Google Doc files in a Google Drive folder to various formats')

  parser.add_argument('drive_path', type=str, help='Folder path on Google Drive for the document source')
  parser.add_argument('-f', '--format', type=str, default='pdf', choices=export_types.keys(), help='Format to export the document as (default=pdf)')
  parser.add_argument('-o', '--out', type=str, default='.', help='Directory to store the exported output (default=current dir)')

  flags = parser.parse_args()
except ImportError:
  flags = None

logger.setLevel(flags.logging_level.upper())

def get_credentials():
  """Gets valid user credentials from storage.

  If nothing has been stored, or if the stored credentials are invalid,
  the OAuth2 flow is completed to obtain the new credentials.

  Returns
    Credentials, the obtained credential.
  """
  home_dir = os.path.expanduser('~')
  credential_dir = os.path.join(home_dir, '.credentials')
  if not os.path.exists(credential_dir):
    os.makedirs(credential_dir)

  credential_path = os.path.join(credential_dir, 'drive-exporter.json')

  store = oauth2client.file.Storage(credential_path)
  credentials = store.get()

  if not credentials or credentials.invalid:
    # See https://developers.google.com/api-client-library/python/auth/installed-app for
    # instructions on how to generate and handle the CLIENT_SECRET_FILE
    flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
    flow.user_agent = APPLICATION_NAME

    credentials = tools.run_flow(flow, store, flags)

    logging.info('Storing credentials to ' + credential_path)
  return credentials

def find_folder(name, parents=[]):
  folder_ids = []
  query = "trashed = false and mimeType = 'application/vnd.google-apps.folder' and name = '%s'" % (name)

  if len(parents) > 0:
    parent_predicate = "'%s' in parents" % parents[0]
    for p in parents[1:]:
      parent_predicate.append(" or '%s' in parents" % (p))

    query = "%s and (%s)" % (query, parent_predicate)

  logging.debug("Getting folders matching query '%s'" % (query))
  resp = svc.files().list(
      q = query,
      fields = "files(id, name)"
    ).execute()

  for f in resp.get('files', []):
    folder_ids.append(f.get('id'))

  if len(folder_ids) < 1:
    msg = "Folder %s not found" % (name)

    for p in parents:
      # TODO add parent paths to error message
      pass

    raise RuntimeError(msg)

  logging.debug("Folder '%s' has ID(s) '%s'" % (name, folder_ids))
  return folder_ids

##################################################
# Must use oAuth2 for API authentication, keys don't work
credentials = get_credentials()
http = credentials.authorize(httplib2.Http())
svc = discovery.build('drive', 'v3', http=http)

folders = flags.drive_path.split('/')
parent_folder_ids = find_folder(folders[0])

for f in folders[1:]:
  parent_folder_ids = find_folder(f, parent_folder_ids)

page_token = None
query = "trashed = false and mimeType = 'application/vnd.google-apps.document' and '%s' in parents" % (parent_folder_ids[0])
while True:
  logging.debug("Getting file list matching query '%s'" % (query))
  resp = svc.files().list(
      q = query,
      fields = "nextPageToken, files(id, name)",
      pageToken = page_token
    ).execute()

  for file in resp.get('files', []):
    file_sufx = flags.format.lower()
    out_path  = os.path.join(flags.out, "%s.%s" % (file.get('name').replace(os.sep, '_'), file_sufx))

    logging.info("Exporting '%s' to '%s'" % (file.get('name'), out_path))
    data = svc.files().export(fileId=file.get('id'), mimeType=export_types[file_sufx]).execute()

    if not os.path.exists(flags.out):
      os.makedirs(flags.out)

    with open(out_path, "wb") as f:
      try:
        f.write(data)
      finally:
        f.close()

  page_token = resp.get('nextPageToken', None)
  if page_token is None:
    break
