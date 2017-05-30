from __future__ import print_function
import httplib2
import os
import datetime

import csv
import json

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/admin-directory_v1-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/admin.directory.group.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Directory API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'admin-directory_v1-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('%s: Storing credentials to ' + credential_path % (datetime.datetime.now()))
    return credentials


def get_domain_groups(authentication):
    service = discovery.build('admin', 'directory_v1', http=authentication)

    print('%s: Getting all email groups in the domain' % (datetime.datetime.now()))
    results = service.groups().list(domain='preventure.com', maxResults=200).execute()
    groups = results.get('groups', [])
    results_object = []

    if not groups:
        print('%s: No groups in the domain.' % (datetime.datetime.now()))
    else:
        print("%s: Detected %s email groups." % (datetime.datetime.now(), len(groups)))
        for group in groups:
            print("%s: >>>Email: %s, Size: %s" % (datetime.datetime.now(), group.get('email'), group.get("directMembersCount")))
            group_members = get_group_members(authentication, group.get('email'))
            results_object.append([group.get('name'), group.get('email'), group.get('description'), {'size': group.get("directMembersCount")}, {'aliases': group.get("aliases")}, group_members])
            break
    return results_object


def get_group_members(authentication, group_email):
    service = discovery.build('admin', 'directory_v1', http=authentication)

    results = service.members().list(groupKey=group_email, maxResults=200).execute()
    members = results.get('members', [])

    result = []
    if not members:
        print('%s: No members in the email group: %s' % (datetime.datetime.now(), group_email))
        return []
    else:
        if group_email:
            for member in members:
                try:
                    del member['etag']
                except KeyError:
                    pass
                result.append((member.get('email'), member.get('role'), member.get('status')))

    return results


def export_data_to_csv(data):
    try:
        with open('groups.csv', 'w', newline='') as csvfile:
            groupwriter = csv.writer(csvfile, delimiter=',',
                                    quotechar='"', quoting=csv.QUOTE_MINIMAL)
            groupwriter.writerow(['name', 'email', 'description', 'size', 'aliases'])
            for group in data:
                print(group['name'], group['email'], group['description'])
                try:
                    groupwriter.writerow([group['name'], group['email'], group['description'], group["directMembersCount"], group['aliases']])
                except KeyError:
                    groupwriter.writerow([group['name'], group['email'], group['description'], group["directMembersCount"]])
    except PermissionError:
        print("Cannot access csv file.")
        raise


def main():
    """Shows basic usage of the Google Admin SDK Directory API.

    Creates a Google Admin SDK API service object and outputs a list of first
    10 users in the domain.
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('admin', 'directory_v1', http=http)

    results = get_domain_groups(http)
    print("%s: Done" % (datetime.datetime.now()))
    filename = 'result.json'
    print("%s: Saving to: %s" % (datetime.datetime.now(), filename))
    with open(filename, 'w') as fp:
        json_file = json.dump(results, fp)

    with open('mycsvfile.csv', 'w') as f:  # Just use 'w' mode in 3.x
        w = csv.DictWriter(f, results.keys())
        w.writeheader()
        w.writerow(results)

if __name__ == '__main__':
    main()

    #TODO: Add instructions for Oauth flow
    #TODO: Add instructions for adding Script to G Suite API Console
    #TODO: Push output files to output folder
    #TODO: Allow CLI input for I/O paths and filenames?
    #TODO: Test script to see if it supports pagination for extra-large domains.