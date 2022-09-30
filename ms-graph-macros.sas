/* --------------------------------------------------------------------------
 Macros for managing the access tokens for the  MS Graph API. 
 Also helpful macros for discovering,  downloading/reading, and uploading 
 file content to OneDrive and SharePoint Online.

 Authors: Joseph Henry, SAS
          Chris Hemedinger, SAS
 Copyright 2022, SAS Institute Inc.

See: 
 https://blogs.sas.com/content/sasdummy/sas-programming-office-365-onedrive
----------------------------------------------------------------------------*/

%let msgraphApiBase = https://graph.microsoft.com/v1.0;
%let msloginBase    = https://login.microsoft.com;

/* Reliable way to check whether a macro value is empty/blank */
%macro isBlank(param);
  %sysevalf(%superq(param)=,boolean)
%mend;

/* We need this function for large file uploads, to telegraph */
/* the file size in the API.                                   */
/* Get the file size of a local file in bytes.                */
%macro getFileSize(localFile=);
  %local rc fid fidc;
  %local File_Size;
  %let rc=%sysfunc(filename(_lfile,&localFile));
  %let fid=%sysfunc(fopen(&_lfile));
  %let File_Size=%sysfunc(finfo(&fid,File Size (bytes)));
  %let fidc=%sysfunc(fclose(&fid));
  %let rc=%sysfunc(filename(_lfile));
  %sysevalf(&File_Size.)
%mend;

/*
  Set the variables that will be needed through the code
  We'll need these for authorization and also for runtime 
  use of the service.
 
  Reading these from a config.json file so that the values
  are easy to adapt for different users or projects.

  Usage:
    %initConfig(configPath=/path-to-your-config-folder);

  configPath should contain the config.json for your app.
  This path will also contain token.json once it's generated
  by the authentication steps.
*/
%macro initConfig(configPath=);
  %global config_root;
  %let config_root=&configPath.;
  filename config "&configPath./config.json";
  %if (%sysfunc(fexist(config))) %then %do;
    libname config json fileref=config;
    data _null_;
     set config.root;
     call symputx('tenant_id',tenant_id,'G');
     call symputx('client_id',client_id,'G');
     call symputx('redirect_uri',redirect_uri,'G');
     call symputx('resource',resource,'G');
    run;
    libname config clear;
    filename config clear;
  %end;
  %else %do;
    %put ERROR: You must create the config.json file in your configPath.; 
    %put The file contents should be:;
    %put   {;
    %put 	  "tenant_id": "your-azure-tenant",;
    %put 	  "client_id": "your-app-client-id",;
    %put 	  "redirect_uri": "https://login.microsoftonline.com/common/oauth2/nativeclient",;
    %put 	  "resource" : "https://graph.microsoft.com";
    %put   };
  %end;
%mend;

/*
  Generate a URL that you will use to obtain an authentication code in your browser window.
  Usage:
   %initConfig(configPath=/path-to-config.json);
   %generateAuthUrl();
*/
%macro generateAuthUrl();
  %if %symexist(tenant_id) %then
    %do;
      /* Run this line to build the authorization URL */
      %let authorize_url=https://login.microsoftonline.com/&tenant_id./oauth2/authorize?client_id=&client_id.%nrstr(&response_type)=code%nrstr(&redirect_uri)=&redirect_uri.%nrstr(&resource)=&resource.;
      %let _currLS = %sysfunc(getoption(linesize));

      /* LS=MAX so URL will not have line breaks for easier copy/paste */
      options nosource ls=max;
      %put Paste this URL into your web browser:;
      %put -- START -------;
      %put &authorize_url;
      %put ---END ---------;
      options source ls=&_currLS.;
    %end;
  %else
    %do;
      %put ERROR: You must use the initConfig macro first.;
    %end;
%mend;

/*
  Utility macro to process the JSON token 
  file that was created at authorization time.
  This will fetch the access token, refresh token,
  and expiration datetime for the token so we know
  if we need to refresh it.
*/
%macro read_token_file(file);
  libname oauth json fileref=&file.;

  data _null_;
    set oauth.root;
    call symputx('access_token', access_token,'G');
    call symputx('refresh_token', refresh_token,'G');
    /* convert epoch value to SAS datetime */
    call symputx('expires_on',(input(expires_on,best32.)+'01jan1970:00:00'dt),'G');
  run;

  libname oauth clear;
%mend;

/*
  Utility macro that retrieves the initial access token
  by redeeming the authorization code that you're granted
  during the interactive step using a web browser
  while signed into your Microsoft OneDrive / Azure account.

  This step also creates the initial token.json that will be
  used on subsequent steps/sessions to redeem a refresh token.
*/
%macro get_access_token(auth_code, debug=0);

  filename token "&config_root./token.json"; 

  proc http url="&msloginBase./&tenant_id./oauth2/token"
    method="POST"
    in="%nrstr(&client_id)=&client_id.%nrstr(&code)=&code.%nrstr(&redirect_uri)=&redirect_uri%nrstr(&grant_type)=authorization_code%nrstr(&resource)=&resource."
    out=token;
    %if &debug>=0 %then
      %do;
        debug level=&debug.;
      %end;
    %else %if &_DEBUG_. ge 1 %then
      %do;
        debug level=&_DEBUG_.;
      %end;
  run;

  %if (&SYS_PROCHTTP_STATUS_CODE. = 200) %then %do;
    %read_token_file(token);
  %end;
  %else %do; 
   %put ERROR: &sysmacroname. failed: HTTP result - &SYS_PROCHTTP_STATUS_CODE. &SYS_PROCHTTP_STATUS_PHRASE.; 
  %end;

  filename token clear;

%mend;

/*
  Utility macro to redeem the refresh token 
  and get a new access token for use in subsequent
  calls to the MS Graph API service.
*/
%macro refresh_access_token(debug=0);

  filename token "&config_root./token.json"; 

  proc http url="&msloginbase./&tenant_id./oauth2/token"
    method="POST"
    in="%nrstr(&client_id)=&client_id.%nrstr(&refresh_token=)&refresh_token%nrstr(&redirect_uri)=&redirect_uri.%nrstr(&grant_type)=refresh_token%nrstr(&resource)=&resource."
    out=token;
    %if &debug. ge 0 %then
      %do;
        debug level=&debug.;
      %end;
    %else %if %symexist(_DEBUG_) AND &_DEBUG_. ge 1 %then
      %do;
        debug level=&_DEBUG_.;
      %end;
  run;

  %if (&SYS_PROCHTTP_STATUS_CODE. = 200) %then %do;
    %read_token_file(token);
  %end;
  %else %do; 
   %put ERROR: &sysmacroname. failed: HTTP result - &SYS_PROCHTTP_STATUS_CODE. &SYS_PROCHTTP_STATUS_PHRASE.; 
  %end;

  filename token clear;

%mend;


/* 
 Use the token information to refresh and gain an access token for this session 
 Usage:
   %initSessionMS365;

 Assumes you have already defined config.json and token.json with
 the authentication steps, and set the config path with %initConfig.
*/

%macro initSessionMS365;

  %if (%isBlank(&config_root.)) %then %do; 
    %put You must use initConfig first to set the configPath;
    %return;
  %end;
  /*
    Our json file that contains the oauth token information
  */
  filename token "&config_root./token.json";

  %if (%sysfunc(fexist(token)) eq 0) %then %do;
   %put ERROR: &config_root./token.json not found.  Run the setup steps to create the API tokens.;
  %end;
  %else %do;
    /*
    If the access_token expires, we can just use the refresh token to get a new one.

    Some reasons the token (and refresh token) might not work:
      - Explicitly revoked by the app developer or admin
      - Password change in the user account for Microsoft Office 365
      - Time limit expiration

    Basically from this point on, user interaction is not needed.

    We assume that the token will only need to be refreshed once per session, 
    and right at the beginning of the session. 

    If a long running session is needed (>3600 seconds), 
    then check API calls for a 401 return code
    and call %refresh_access_token if needed.
    */

    %read_token_file(token);

    filename token clear;

    /* If this is first use for the session, we'll likely need to refresh  */
    /* the token.  This will also call read_token_file again and update */
    /* our token.json file.                                                */
    %refresh_access_token();
  %end;  
%mend;

/* For SharePoint Online, list the main document libraries in the root of a SharePoint site */
/* Using the /sites methods in the Microsoft Graph API            */
/* May require the Sites.ReadWrite.All permission for your app    */
/* See https://docs.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0 */
/* Set these values per your SharePoint Online site.
   Ex: https://yourcompany.sharepoint.com/sites/YourSite 
    breaks down to:
       yourcompany.sharepoint.com -> hostname
       /sites/YourSite -> sitepath

   This example uses the /drive method to access the files on the
   Sharepoint site -- works just like OneDrive.
   API also supports a /lists method for SharePoint lists.
   Use the Graph Explorer app to find the correct APIs for your purpose.
    https://developer.microsoft.com/en-us/graph/graph-explorer

  Usage:
    %listSiteLibraries(siteHost=yoursite.company.com,
          sitePath=/sites/YourSite,
          out=work.OutputListData);
*/
%macro listSiteLibraries(siteHost=,sitePath=,out=work.siteLibraries);
  filename resp TEMP;
  proc http url="&msgraphApiBase./sites/&siteHost.:&sitepath.:/drive"
       oauth_bearer="&access_token"
       out = resp;
  	 run;
  %if (&SYS_PROCHTTP_STATUS_CODE. = 200) %then %do;
    libname jresp json fileref=resp;
    data &out.;
     set jresp.root(drop=ordinal:);
    run;
    libname jresp clear;
  %end;
  %else %do; 
   %put ERROR: &sysmacroname. failed: HTTP result - &SYS_PROCHTTP_STATUS_CODE. &SYS_PROCHTTP_STATUS_PHRASE.; 
  %end;

  filename resp clear;
%mend;

/* 
 For OneDrive, fetch the list of Drives available to the current user.
 
 Output is a data set with the list of available Drives and IDs, for use in later 
 routines.

 This creates a data set with the one record for each drive.
 Note that even if you think you have just one drive, the system
 might track others behind-the-scenes.

 Usage:
   %listMyDrives(out=work.DriveData);
*/
%macro listMyDrives(out=work.drives);
  filename resp TEMP;
  proc http url="&msgraphApiBase./me/drives/"
       oauth_bearer="&access_token"
       out = resp;
  	 run;

  %if (&SYS_PROCHTTP_STATUS_CODE. = 200) %then %do;
    libname jresp json fileref=resp;

    proc sql;
      create table &out. as 
        select t1.id, 
          t1.name, 
          scan(t1.webUrl,-1,'/') as driveDisplayName,
          t1.createdDateTime,
          t1.description,
          t1.driveType,
          t1.lastModifiedDateTime,
          t2.displayName as lastModifiedName,
          t2.email as lastModifiedEmail,
          t2.id as lastModifiedId,
          t1.webUrl
        from jresp.value t1 inner join jresp.lastmodifiedby_user t2 on 
           (t1.ordinal_value=t2.ordinal_lastModifiedBy);
    quit;
    libname jresp clear;
  %end;
  %else %do; 
   %put ERROR: &sysmacroname. failed: HTTP result - &SYS_PROCHTTP_STATUS_CODE. &SYS_PROCHTTP_STATUS_PHRASE.; 
  %end;
  filename resp clear;
%mend;

/*
 List items in a folder in OneDrive or SharePoint
 The Microsoft Graph API returns maximum 200 items, so if the collection
 contains more we need to iterate through a list.

 The API response contains a URL endpoint to fetch the next
 batch of items, if there is one.

 Use folderId=root to list the root items of the "Drive" (OneDrive or SharePoint library),
 else use the folder ID of the folder you discovered in a previous call.
*/
%macro listFolderItems(driveId=, folderId=root, out=work.folderItems); 

  %local nextLink batchnum;

  /* endpoint for initial list of items */
  %let nextLink = &msgraphApiBase./me/drives/&driveId./items/&folderId./children;
  %let batchnum = 1;
  data _folderItems0;
   length name $ 500;
   stop;
  run;

  %do %until (%isBlank(%str(&nextLink)));
    filename resp TEMP;
    proc http url="&nextLink."
         oauth_bearer="&access_token"
         out = resp;
    	 run;
     
    libname jresp json fileref=resp; 

    /* holding area for attributes that might not exist */
    data _value;
      length name $ 500   
      size  8   
      webUrl $ 500   
      lastModifiedDateTime $ 20   
      createdDateTime $ 20   
      id $ 50   
      eTag $ 50   
      cTag $ 50   
      _microsoft_graph_downloadUrl $ 2000   
      fileMimeType $ 75   
      isFolder  8   
      folderItemsCount  8;   
      %if %sysfunc(exist(JRESP.VALUE)) %then
        %do;
          set JRESP.VALUE;
        %end;
    run;

    data _value_file;
      length ordinal_value 8 mimeType $ 75 ;
      %if %sysfunc(exist(JRESP.VALUE_FILE)) %then %do;
        set JRESP.VALUE_FILE;
      %end;
    run;

    data _value_folder;
      length ordinal_value 8 ordinal_folder 8 childCount 8;
      %if %sysfunc(exist(JRESP.VALUE_FOLDER)) %then %do;
        set JRESP.VALUE_FOLDER;
      %end;
    run;

    proc sql;
      create table _folderItems&batchnum. as 
        select t1.name, t1.size, t1.webUrl length=500,
          t1.lastModifiedDateTime,
          t1.createdDateTime,
          t1.id,
          t1.eTag,
          t1.cTag,
          t1._microsoft_graph_downloadUrl,
          t3.mimeType as fileMimeType,
        case 
          when t2.ordinal_folder is missing then 0
          else 1
        end 
      as isFolder,
        t2.childCount as folderItemsCount
      from _value t1 left join _value_folder t2 
        on (t1.ordinal_value=t2.ordinal_folder)
      left join _value_file t3 on (t1.ordinal_value=t3.ordinal_value)
      ;
    quit;

    /* clear placeholder attributes */
    proc delete data=work._value_folder work._value_file work._value ; run;

     %put NOTE: Batch &batchnum: Gathered &sysnobs. items;
    /* check for a next link for more entries */
    %let nextLink=;
    data _null_;
     set jresp.alldata(where=(p1='@odata.nextLink'));
     call symputx('nextLink',value);  
    run;
    %let batchnum = %sysevalf(&batchnum. + 1);

    libname jresp clear;
    filename resp clear;
  %end;
  
  data &out;
   set _folderItems:;
  run;

  proc datasets nodetails nolist;
   delete _folderItems:;
  run;

%mend;

/* Download a OneDrive or SharePoint file                        */
/* Each file has a specific download URL that works with the API */
/* This macro routine finds that URL and use PROC HTTP to GET    */
/* the content and place it in the local destination path        */
%macro downloadFile(driveId=,folderId=,sourceFilename=,destinationPath=);
  %local driveId folderId dlUrl _opt;
  %let _opt = %sysfunc(getoption(quotelenmax)); 
  options noquotelenmax;

  %listFolderItems(driveId=&driveId., folderId=&folderId., out=__tmpLst);

  proc sql noprint;
    select _microsoft_graph_downloadUrl into: dlUrl from folderItems
      where name="&sourceFilename";
  quit;

  proc delete data=work.__tmpLst; run;

  %if %isBlank(&dlUrl) %then %do;
    %put ERROR: No file named &sourceFilename. found in folder.;
  %end;
  %else %do;
    filename dlout "&destinationPath./&sourceFilename.";

    proc http url="%nrbquote(&dlUrl.)"
      oauth_bearer="&access_token"
      out = dlOut;
    run;

    %put NOTE: Download file HTTP result - &SYS_PROCHTTP_STATUS_CODE. &SYS_PROCHTTP_STATUS_PHRASE.; 

    %if (&SYS_PROCHTTP_STATUS_CODE. = 200) %then %do;
      %put NOTE: File downloaded to &destinationPath./&sourceFilename., %getFilesize(localFile=&destinationPath./&sourceFilename) bytes;
    %end;
    %else %do;
     %put WARNING: Download file NOT successful.;
    %end;

    filename dlout clear;
  %end;
  options &_opt;
%mend;


/* 
   Use an UploadSession in the Microsoft Graph API to upload a file.   

   This can handle large files, greater than the 4MB limit used by     
   PUT to the :/content endpoint.                                       
   The Graph API doc says you need to split the file into chunks, but  
   using EXPECT_100_CONTINUE in PROC HTTP seems to do the trick.       

   We do need to know the total file size in bytes before using the API, so
   this code includes a file-size check.

   Credit to Muzzammil Nakhuda at SAS for figuring this out.           

   Usage:
    %uploadFile(driveId=&driveId.,folderId=&folder.,
       sourcePath=<local-SAS-folder-where-file-is>,
       sourceFilename=<local-SAS-file-name>);
*/
%macro uploadFile(driveId=,folderId=,sourcePath=,sourceFilename=) ;
  %local driveId folderId fileSize _opt uploadURL;
  %let _opt = %sysfunc(getoption(quotelenmax)); 
  options noquotelenmax;
  filename resp temp;
 
   /* Create an upload session to upload the file.                                                */
   /* If a file of the same name exists, we will REPLACE it.                                      */
   /* The API doc says this should be POST, but since we provide a body with conflict directives, */
   /* it seems we must use PUT.                                                                   */
   proc http url="&msgraphApiBase./me/drives/&driveId./items/&folderId.:/&sourceFilename.:/createUploadSession"
     method="PUT"
     in='{ "item": {"@microsoft.graph.conflictBehavior": "replace" } }'
     out=resp
     oauth_bearer="&access_token";
   run;
    %put NOTE: Create Upload Session: HTTP result - &SYS_PROCHTTP_STATUS_CODE. &SYS_PROCHTTP_STATUS_PHRASE.; 

    %if (&SYS_PROCHTTP_STATUS_CODE. = 200) %then %do;
      libname resp JSON fileref=resp;   
      data _null_;
      set resp.root; 
       call symputx('uploadURL',uploadUrl);
      run;
   
      filename _tosave "&sourcePath./&sourceFilename.";
   
      %let fileSize=%getFileSize(localfile=&sourcePath./&sourceFilename.);  
      %let end_index=%eval(&fileSize.-1);
   
      %put Uploading &sourcePath./&sourceFilename., file size of &fileSize bytes.;
   
      filename hdrout temp;
      filename resp temp;
   
      proc http url= "&uploadURL."
         method="PUT"
         in=_tosave
         out=resp
         oauth_bearer="&access_token"
         headerout=hdrout
         EXPECT_100_CONTINUE 
         ;
         headers
           "Content-Range"="bytes 0-&end_index/&fileSize."
           ;
       run;
    
       %put NOTE: Upload file HTTP result - &SYS_PROCHTTP_STATUS_CODE. &SYS_PROCHTTP_STATUS_PHRASE.; 
       /* HTTP 200 if success, 201 if new file was created */
       %if (%sysfunc(substr(&SYS_PROCHTTP_STATUS_CODE.,1,1))=2) %then %do;
         /*
           Successful call returns a json response that describes the item uploaded.
           This step pulls out the main file attributes from that response.
         */
         libname attrs json fileref=resp;
         %if %sysfunc(exist(attrs.root)) %then %do;
           data _null_;
            length filename $ 100 createdDate 8 modifiedDate 8 filesize 8;
            set attrs.root;
            filename = name;
            modifiedDate = input(lastModifiedDateTime,anydtdtm.);
            createdDate  = input(createdDateTime,anydtdtm.);
            format createdDate datetime20. modifiedDate datetime20.;
            filesize = size;
            msg = cat(name,' uploaded. File size: ',put(size,comma15. -L),' bytes. Created ',put(createdDate,datetime20.), ', Modified ',put(modifiedDate,datetime20.));
            put 'NOTE: ' msg;
           run;
         %end;
         libname attrs clear;
       %end;
       %else %do;
         %put WARNING: File upload failed!;
         %if (%sysfunc(fexist(resp))) %then %do;
          data _null_; rc=jsonpp('resp','log'); run;
         %end;
        %if (%sysfunc(fexist(hdrout))) %then %do;
          data _null_; infile hdrout; input; put _infile_; run;
        %end;
       %end;
       filename _tosave clear;    
       filename hdrout clear;
     %end;
     %else %do;
      %put WARNING: Upload session not created!; 
      %if (%sysfunc(fexist(resp))) %then %do;
        data _null_; rc=jsonpp('resp','log'); run;
      %end;
     %end;
     filename resp clear;
     options &_opt;
 %mend;
