import requests
import json
import msal
import urllib
import os

"""
Class that handles uploading, downloading, and deleting sharepoint files using the OAuth2 authentication and the Microsoft Graph API

Michael P. Vossen
10/19/2022


"""



class OAuth2_SharePoint:
    def __init__(self, First=False):
        """
        The initalization of the object. 
        
        Input:
            First (bool)
                A boolean for if you are running the object for the first time after you have changed
                the application id and secret value.  If set to True, this will run the interactive mode
                that allows you to approve this application to use your credentials.
        """
        
        
        """
        The section below contains the OAuth2 related secrets
        Note: These values expire every 12 months.  When you need to update
        these values contact Adam Zembrosky (asz@uwm.edu) and run this class
        with the input First set to true. 
        
        Last Updated: 10/17/2022
        """
        
 
        self.__application_id__ = 
        self.__secret_value__ = 
            

        
        
        """
        This section contains the authentication for the Innovative Weather
        service account.  We must always use the Innovative Weather service
        account.
        """
        
        self.__uwm_id__ = 
        self.__uwm_password__ = 
        
        
        """
        This section conatins the UWM tennant ID so OAuth2 recognizes we are 
        associated with UWM.
        
        This comes from https://login.microsoftonline.com/panthers.onmicrosoft.com/.well-known/openid-configuration
        Its the long string of numbers and letters between .com/ and /oauth2 in the top url.
        """
        
        self.__tennant_id__ = "0bca7ac3-fcb6-4efd-89eb-6de97603cf21"
        
        self.__authority_url__ =  f'https://login.microsoftonline.com/{self.__tennant_id__}'
        
        
        """
        This section sets the premissions we need to accomplish our goals in 
        SharePoint
        """
        self.permissions = ['Sites.ReadWrite.All', 'Files.ReadWrite.All']
        
        """
        This sets the general url of UWM's SharePoint
        """
        self.sharepoint_host_name = 'panthers.sharepoint.com'
        
        
        """
        This is the name of our SharePoint site
        """

        self.sharepoint_name = "InnovativeWeatherWebsite-Group"
        
        

        
        """
        Finally we are going to set the resorce that we are going to access.
        Version is set to v1.0 since we are using Graph Version 1
        """
        
        
        self.resource_url = 'https://graph.microsoft.com/'
        self.resource_version = "v1.0"
        
        self.endpoint = self.resource_url + self.resource_version
        
        
        """
        Now we need to go to Microsoft and ask for a tocken using the credentials
        we inputed above.  Microsoft knows that we have permission to retreive this 
        token.  Tokens are used like a username and password for this program to 
        access SharePoint
        
        """
        if First == False:
            pca = msal.ConfidentialClientApplication(self.__application_id__, self.__secret_value__, authority = self.__authority_url__)
            
            #get our oauth token            
            self.__token__ = pca.acquire_token_by_username_password(self.__uwm_id__, self.__uwm_password__, self.permissions)
            
            #create the header with our oauth token for later api requests.
            self.__headers__ = {'Authorization': 'Bearer {}'.format(self.__token__['access_token'])}
            
        
            """
            When using this method everything we access has an ID.  We will need 
            to use the ID to access the data we want.  So below I get the IDs
            for the SharePoint site and the Documents drive where our files are
            located.
            """
        
            result = requests.get(f'{self.endpoint}/sites/{self.sharepoint_host_name}:/sites/{self.sharepoint_name}', headers=self.__headers__)
            site_info = result.json()
            self.__site_id__ = site_info['id']
            result = requests.get(f'{self.endpoint}/sites/{self.__site_id__}/drive', headers=self.__headers__)
            drive_info = result.json()
            self.__drive_id__ = drive_info['id']
        else:
            """
            Now we need to go to Microsoft and ask for a tocken using the credentials
            we inputed above.  Microsoft knows that we have permission to retreive this 
            token.  Tokens are used like a username and password for this program to 
            access SharePoint
            
            """
            
            pca = msal.PublicClientApplication(self.__application_id__, authority = self.__authority_url__)
            
            pca.acquire_token_interactive(self.permissions)

        
    
    def upload_file(self, local_item, item_remote_directory, multiple = False):
        """
        This method uses the Microsoft Graph API to upload small files to a SharePoint 
        destintation.  
        
        Microsoft Graph API Upload Small File Documentation:
        https://learn.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0&tabs=http

        Parameters
        ----------
        local_item : STRING
            Location and name of an item on the local computer that you want to upload to SharePoint
        item_remote_directory : STRING
            Location on SharePoint where the file should be uploaded
        multiple : BOOL, optional
            Whether this is being ran on multiple files (True) or a single file (False). The default is False.

        Returns
        -------
        result (STRING (json like))
            The result return from Microsft Graph after uploading the file.

        """
        
        
        
        #Get size to determine if the file is a large file.  I convert it to MB.  If files are larger than 2 MB we
        #need to upload the file in chuncks.
        size = os.path.getsize(local_item) / (1 * (10 ** 6))
        
        #if file is greater than 2MB use the chunk method
        if size > 2:
            result = self.upload_large_file(local_item, item_remote_directory, multiple=multiple)
            #return json with the upload results.  This is mainly for debugging.
            return result
        #if file is less than 2MB we can upload it all at once
        else:
            
            
            if multiple == False:
                self.check_location(item_remote_directory)
                
            #split the file path so we have the file path seperate from the file name    
            item_local_directory, item_name = self.seperate_path_file(local_item)
            
            #if the file path is missing a / at the end, add it
            if item_remote_directory[-1] != "/":
                item_remote_directory += "/"
            
            
            #this is the actual communication with the GraphAPI
            
            #get the file name on SharePoint
            file_rel_path = urllib.parse.quote(item_name)
            #get the directory path on SharePoint
            folder_rel_path = urllib.parse.quote(item_remote_directory)
            #get the folder id from sharepoint
            result = requests.get(f'{self.endpoint}/drives/{self.__drive_id__}/root:/{folder_rel_path}', headers=self.__headers__)
            #the result from above is a json with a lot of information.  We just need the id
            folder_id = result.json()['id']
            
            #upload the file to sharepoint
            result = requests.put(f'{self.endpoint}/drives/{self.__drive_id__}/items/{folder_id}:/{file_rel_path}:/content'
                              ,headers = self.__headers__
                              ,data = open(f"{item_local_directory}{item_name}", 'rb').read())
            #return a json with the result of the upload.  This is just for debugging purposes
            return result
            
        
    def download_file(self, remote_item, item_local_directory):
        """
        This method uses the Microsoft Graph API to download files from a SharePoint 
        location.  
        
        Microsoft Graph API Download File Documentation:
        https://learn.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0&tabs=http

        Parameters
        ----------
        remote_item : STRING
            Item location and name on SharePoint that you wish to download
        item_local_directory : STRING
            Location on the local computer where the file will download to


        Returns
        -------
        None.

        """

        #split the file path and file name.
        item_remote_directory, item_name = self.seperate_path_file(remote_item)
        
        #if the file path is missing a / add it to the end of the filepath
        if item_local_directory[-1] != "/":
            item_local_directory += "/"
            

        #the full file path
        file_path = f"{item_remote_directory}{item_name}"
        
        #get the file name to build the url for SharePoint
        file_url = urllib.parse.quote(file_path)
        
        #get the id of the directory on SharePoint we are getting the file from
        result = requests.get(f'{self.endpoint}/drives/{self.__drive_id__}/root:/{file_url}', headers = self.__headers__)
        file_info = result.json()
        file_id = file_info['id']
        
        #get the file off of SharePoint
        result = requests.get(f'{self.endpoint}/drives/{self.__drive_id__}/items/{file_id}/content', headers = self.__headers__)
        open(f"{item_local_directory}{item_name}", 'wb').write(result.content)
        
        
        
    def create_folder(self, folder_loc, folder_name):
        """
        This method uses the Microsoft Graph API to create a new folder in SharePoint.
        
        
        Microsoft Graph API Create Folder Documentation:
        https://learn.microsoft.com/en-us/graph/api/driveitem-post-children?view=graph-rest-1.0&tabs=http

        Parameters
        ----------
        folder_loc : STRING
            SharePoint location where the new folder will be created
        folder_name : STRING
            Name of the new folder

        Returns
        -------
        None.

        """
        
        #get the folder url name of the folder you want your new folder to be a part of
        file_url = urllib.parse.quote(folder_loc)
        #get the id of the folder that your new folder is going to be a subfolder of.
        result = requests.get(f'{self.endpoint}/drives/{self.__drive_id__}/root:/{file_url}', headers = self.__headers__)
        folder_id = result.json()['id']
        
        #This time I have to pass a json to the api.  So, I need to add to the header with the oauth token that 
        #I'm passing in a json
        header = self.__headers__
        header["Content-Type"] = "application/json"
        
        #create the folder on SharePoint
        requests.post(f'{self.endpoint}/drives/{self.__drive_id__}/items/{folder_id}/children',headers = header,data = json.dumps({"name" : folder_name, "folder": { }}))
        
    
    def delete_item(self, item):
        """
        This method uses the Microsoft Graph API to delete files and folder on 
        SharePoint.
        
        
        Microsoft Graph API Create Folder Documentation:
        https://learn.microsoft.com/en-us/graph/api/driveitem-delete?view=graph-rest-1.0&tabs=http

        Parameters
        ----------
        item : STRING
            Full path to file or folder on SharePoint

        Returns
        -------
        None.

        """
        
        #get the url name of the file that you want to delete
        file_url = urllib.parse.quote(item)
        #get the id of the file you want to delete
        result = requests.get(f'{self.endpoint}/drives/{self.__drive_id__}/root:/{file_url}', headers = self.__headers__)
        item_id = result.json()['id']
        
        #delete the file
        requests.delete(f"{self.endpoint}/drives/{self.__drive_id__}/items/{item_id}", headers = self.__headers__)
        
        
    def list_folder(self, remote_location):
        """
        This method uses the Microsoft Graph API to list files and folder in 
        SharePoint.
        
                
        Microsoft Graph API Create Folder Documentation:
        https://learn.microsoft.com/en-us/graph/api/driveitem-list-children?view=graph-rest-1.0&tabs=http

        Parameters
        ----------
        remote_location : STRING
            Full path to the location you wish to list the folder contents

        Returns
        -------
        file_names : LIST
            A list of files with full paths that are located in the folder you specified

        """
        #get the url name for the folder you want to list
        file_url = urllib.parse.quote(remote_location)
        
        #get the id of the folder that you want to list
        result = requests.get(f'{self.endpoint}/drives/{self.__drive_id__}/root:/{file_url}', headers = self.__headers__)
        folder_id = result.json()['id']
        
        #list the files
        output = requests.get(f'{self.endpoint}/drives/{self.__drive_id__}/items/{folder_id}/children', headers = self.__headers__)
        
        #from the previous step we get a json with a lot of information.  We just want file names which are contained in the value key.
        output = output.json()["value"]
        
        #initalize list for the file path and file names
        file_names = []
        
        #for each entry in the values key
        for file in output:
            #get the string name, add the location in front, and append it to the file_names list
            file_names.append(remote_location + "/" + file["name"])
            
            
        #return a list of file names
        return file_names
        
        
    def seperate_path_file(self, full_file):
        """
        A method to seperate the folder path from the file name.

        Parameters
        ----------
        full_file : STRING
            Full path to a file

        Returns
        -------
        folder_path : STRING
            Folder directory location of a file
        file : STRING
            Name of the file

        """
        #if the file name has a / it is a full file path.
        #we then need to split the string by / to find the file name
        if full_file.find("/") == -1 and full_file.find(os.sep) != -1:
            delim = os.sep
        #if no / exists we just have a file name and we need to do nothing more
        elif full_file.find("/") == -1 and full_file.find(os.sep) == -1:
            return "", full_file
        
        #just a fail safe
        else:
            delim = "/"
         
        #split the file path by /
        path_list = full_file.split(delim)
        
        #the last index is the file name.  So I remove it and save it to a seperate variable.
        file = path_list.pop(-1)
        
        #join the list of folders back to a string sepearted by /
        folder_path = "/".join(path_list) + "/"
        
        #return the folder path and filename.
        return folder_path, file

    
    
    def check_location(self, directory):
        """
        Method to make sure the directories specified exist in SharePoint.  If they
        don't this method creates the folder.

        Parameters
        ----------
        directory : STRING
            Full path to a directory.

        Returns
        -------
        None.

        """
        #split the directory path into a list of directories.
        #we have to loop through each one since Graph cannot
        #create directories recursivly
        directory_list = directory.split("/")
        #base directory.  This always exists and is named "Shared Documents"
        start_dir = directory_list[0]
        #if there is more than one directory
        if len(directory_list) > 1:
            #for each directory starting with the first directory in "Shared Documents"
            for i in range(1,len(directory_list)):
                #create the folder in SharePoint. Using our create folder method
                self.create_folder(start_dir, directory_list[i])
                
                #add this directory to the end of the start directory
                start_dir = start_dir + "/" + directory_list[i]
        
    def upload_multiple_files(self, files, remote_location):
        """
        Method to upload multiple files to SharePoint

        Parameters
        ----------
        files : LIST
            List of full paths to files on the local machine that you wish to 
            upload to SharePoint
        remote_location : STRING
            Location on SharePoint where files should be uploaded

        Returns
        -------
        None.

        """
        #check to see if the remote location exists
        self.check_location(remote_location)
        
        #for each file upload it.
        for file in files:
            self.upload_file(file, remote_location, multiple=True)
            
    def download_multiple_files(self, files, local_location):
        """
        Method for downloading multiple files from SharePoint

        Parameters
        ----------
        files : LIST
            List of full file paths of the files on SharePoint you wish to download.
        local_location : STRING
            The directory on the local machine where you wish to download the files to.

        Returns
        -------
        None.

        """
        #for each file download the file
        for file in files:
            self.download_file(file, local_location)
            
            
    def upload_large_file(self, local_item, item_remote_directory, multiple = False):
        """
        This method uses the Microsoft Graph API to upload large files in chunks to a SharePoint 
        destintation.  
        
        Microsoft Graph API Upload Large File Documentation:
        https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0

        Parameters
        ----------
        local_item : STRING
            Location and name of an item on the local computer that you want to upload to SharePoint
        item_remote_directory : STRING
            Location on SharePoint where the file should be uploaded
        multiple : BOOL, optional
            Whether this is being ran on multiple files (True) or a single file (False). The default is False.

        Returns
        -------
        results : LIST
            List of json results of uploading the chunks with Graph.  This is just for debugging.

        """
        
        #check to see if location exists
        if multiple == False:
            self.check_location(item_remote_directory)

        item_local_directory, item_name = self.seperate_path_file(local_item)
            
        #add a / to the end of the directory path if it doesn't exist
        if item_remote_directory[-1] != "/":
            item_remote_directory += "/"
        
        #####################################
        #start by creating the upload session.
        #####################################
        
        #get file path url name
        file_rel_path = urllib.parse.quote(item_name)
        #get folder path url name
        folder_rel_path = urllib.parse.quote(item_remote_directory)
        
        #get id of the folder we are uploading to
        result = requests.get(f'{self.endpoint}/drives/{self.__drive_id__}/root:/{folder_rel_path}', headers=self.__headers__)
        folder_id = result.json()['id']
        
        #create the upload session.
        result = requests.post(f'{self.endpoint}/drives/{self.__drive_id__}/items/{folder_id}:/{file_rel_path}:/createUploadSession'
                          ,headers = self.__headers__
                          ,data = json.dumps({"item" : {"@microsoft.graph.conflictBehavior": "replace", "name" : item_name }}))
        
        #############################
        #upload the individual chunks
        #############################
        
        #get the file size so we can know how to break up the chunks
        size = os.path.getsize(local_item)
        #open the file as bytes
        file_bytes = open(local_item, "rb")
        
        #define the chunk size.  Don't change this.
        chunk = 1024 * 320
        
        #read the first chunk
        byte = file_bytes.read(chunk)
        
        #the starting byte for the chunk
        start = 0
        #the ending byte for the chunk
        end = chunk - 1
        
        #add the content of this chunk to a json to pass to Graph
        result = json.loads(result.text)
        
        #a list of api results.  This is for debugging.
        results = []
        
        #while there is bytes to upload
        while byte:
            #the number of bytes we are uploading.  This needs to be passed to Graph
            num_bytes = len(byte)
            
            #upload the chunk to SharePoint
            chunk_result = requests.put(result["uploadUrl"],headers = {"Content-Length" : str(num_bytes) ,"Content-Range" : f"bytes {str(start)}-{str(end)}/{str(size)}"}, data = byte)
            
            #add the result of the upload to the list
            results.append(json.loads(chunk_result.text))
            
            #read the next chunk
            byte = file_bytes.read(327680)
            
            #check the length
            num_bytes = len(byte)
            
            #change the start and end byte.
            start += chunk
            end += num_bytes
            
        
        #when we are done end the upload session
        delete_result = requests.delete(result["uploadUrl"])
        return results
        
            

        


#the debugging part of the code.
if __name__ == "__main__":
    auth = OAuth2_SharePoint(False)
    result = auth.upload_file("C:/Users/mpvossen/Downloads/WUWM30.wav", "WUWM/")
    #auth.download_file("WUWM/WUWM30.wav", ".")