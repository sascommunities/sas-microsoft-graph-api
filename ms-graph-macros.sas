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

/* Reliable way to check whether a macro value is empty/blank */
%macro isBlank(param);
  %sysevalf(%superq(param)=,boolean)
%mend;

/* Check to see if base URLs for services are      */
/* initialized/overridden.                         */
/* If not, then define them to the common defaults */
%macro initBaseUrls();
 %if %symexist(msloginBase) = 0 %then %do;
   %global msloginBase;
 %end; 
 %if %isBlank(&msloginBase.) %then %do;
   %let msloginBase = https://login.microsoftonline.com;
 %end;
 %if %symexist(msgraphApiBase) = 0 %then %do;
   %global msgraphApiBase;
 %end; 
  %if %isBlank(&msgraphApiBase.) %then %do;
   %let msgraphApiBase = https://graph.microsoft.com/v1.0;
 %end;
%mend;
%initBaseUrls();

/* We need this function for large file uploads, to telegraph */
/* the file size in the API.                                   */
/* Get the file size of a local file in bytes.                */
%macro getFileSize(localFile=);
  %local rc fid fidc;
  %local File_Size;
  %let File_Size = -1;
  %let rc=%sysfunc(filename(_lfile,&localFile));
  %if &rc. = 0 %then %do; 
    %let fid=%sysfunc(fopen(&_lfile));
    %if &fid. > 0 %then %do;
      %let File_Size=%sysfunc(finfo(&fid,File Size (bytes)));
      %let fidc=%sysfunc(fclose(&fid));
    %end;
    %let rc=%sysfunc(filename(_lfile));
  %end;
  %sysevalf(&File_Size.)
%mend;

/*
  Set the variables that will be needed through the code
  We'll need these for authorization and also for runtime 
  use of the service.
 
  Reading these from a config.json file so that the values
  are easy to adapt for different users or projects. The config.json
  can be in a file system or in SAS Content folders (SAS Viya only).

  Usage:
    %initConfig(configPath=/path-to-your-config-folder);
  
  If using SAS Content folders on SAS Viya, specify the content
  folder and SASCONTENT=1.

    %initConfig(configPath=/Users/&_clientuserid/My Folder/.creds,sascontent=1);

  configPath should contain the config.json for your app.
  This path will also contain token.json once it's generated
  by the authentication steps.
*/
%macro initConfig(configPath=,sascontent=0);
  %global config_root m365_usesascontent;
  %let m365_usesascontent = &sascontent.;
  %let config_root=&configPath.;
  %if &m365_usesascontent = 1 %then %do;
    filename config filesrvc 
      folderpath="&configPath."
      filename="config.json";
  %end;
  %else %do;
    filename config "&configPath./config.json";
  %end;
  %put NOTE: Establishing Microsoft 365 config root to &config_root.;
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
    %put 	  "redirect_uri": "&msloginBase./common/oauth2/nativeclient",;
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
      %let authorize_url=&msloginBase./&tenant_id./oauth2/authorize?client_id=&client_id.%nrstr(&response_type)=code%nrstr(&redirect_uri)=&redirect_uri.%nrstr(&resource)=&resource.;
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
  %put M365: Reading token info from %sysfunc(pathname(&file.));
  libname oauth json fileref=&file.;

  data _null_;
    set oauth.root;
    call symputx('access_token', access_token,'G');
    call symputx('refresh_token', refresh_token,'G');
    /* convert epoch value to SAS datetime */
    call symputx('expires_on',(input(expires_on,best32.)+'01jan1970:00:00'dt),'G');
  run;
  %put M365: Token expires on %left(%qsysfunc(putn(%sysevalf(&expires_on.+%sysfunc(tzoneoff() )),datetime20.)));

  libname oauth clear;
%mend;

/* Assign the TOKEN fileref to location that  */
/* depends on whether we're using SAS Content */
%macro assignTokenFileref();
  %if &m365_usesascontent = 1 %then %do;
    filename token filesrvc 
      folderpath="&config_root."
      filename="token.json";
  %end;
  %else %do;
    filename token "&config_root./token.json";
  %end;
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

   %assignTokenFileref();

  proc http url="&msloginBase./&tenant_id./oauth2/token"
    method="POST"
    in="%nrstr(&client_id)=&client_id.%nrstr(&code)=&auth_code.%nrstr(&redirect_uri)=&redirect_uri%nrstr(&grant_type)=authorization_code%nrstr(&resource)=&resource."
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
 
  %put M365: Refreshing access token for M365;
   %assignTokenFileref();

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
   %assignTokenFileref();

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

  %local driveId nextLink batchnum;

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

  /* Use DATA step functions here to escape & to avoid warnings for unresolved symbols */
  data _null_;
    set __tmpLst;
    length resURL $ 2000;
    where name="&sourceFilename";
    resURL = tranwrd(_microsoft_graph_downloadUrl,'&','%str(&)');
    call symputx('dlURL',resURL);
  run;

  proc delete data=work.__tmpLst; run;

  %if %isBlank(&dlUrl) %then %do;
    %put ERROR: No file named &sourceFilename. found in folder.;
  %end;
  %else %do;
    filename dlout "&destinationPath./&sourceFilename.";

    proc http url="&dlUrl."
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
Split a file into same-size chunks, often needed for HTTP uploads
of large files via an API

Sample use:
 %splitFile(sourceFile=c:\temp\register-hypno.gif, 
     maxSize=327680,
     metadataOut=work.chunkMeta,
     chunkLoc=c:\temp\chunks);
*/

%macro splitFile(sourceFile=,
 maxSize=327680,
 metadataOut=,
 /* optional, will default to WORK */
 chunkLoc=);

  %local filesize maxSize numChunks buffsize ;
  %let buffsize = %sysfunc(min(&maxSize,4096));
  %let filesize = %getFileSize(localFile=&sourceFile.);
  %let numChunks = %sysfunc(ceil(%sysevalf( &filesize / &maxSize. )));
  %put NOTE: Splitting &sourceFile. (size of &filesize. bytes) into &numChunks parts;

  %if %isBlank(&chunkLoc.) %then %do;
    %let chunkLoc = %sysfunc(getoption(WORK));
  %end;

  /* This DATA step will do the chunking.                                 */
  /* It's going to read the original file in segments sized to the buffer */
  /* It's going to write that content to new files up to the max size     */
  /* of a "chunk", then it will move on to a new file in the sequence     */
  /* All resulting files should be the size we specified for chunks       */
  /* except for the last one, which will be a remnant                     */
  /* Along the way it will build a data set with the metadata for these   */
  /* chunked files, including the file location and byte range info       */
  /* that will be useful for APIs that need that later on                 */
  data &metadataOut.(keep=original originalsize chunkpath chunksize byterange);
    length 
      filein 8 fileid 8 chunkno 8 currsize 8 buffIn 8 rec $ &buffsize fmtLength 8 outfmt $ 12
      bytescumulative 8
      /* These are the fields we'll store in output data set */
      original $ 250 originalsize 8 chunkpath $ 500 chunksize 8 byterange $ 50;
    original = "&sourceFile";
    originalsize = &filesize.;
    rc = filename('in',"&sourceFile.");
    filein = fopen('in','S',&buffsize.,'B');
    bytescumulative = 0;
    do chunkno = 1 to &numChunks.;
      currsize = 0;
      chunkpath = catt("&chunkLoc./chunk_",put(chunkno,z4.),".dat");
      rc = filename('out',chunkpath);
      fileid = fopen('out','O',&buffsize.,'B');
      do while ( fread(filein)=0 ) ;
        call missing(outfmt, rec);
        rc = fget(filein,rec, &buffsize.);
        buffIn = fcol(filein);
        if (buffIn - &buffsize) = 1 then do;
          currsize + &buffsize;
          fmtLength = &buffsize.;
        end;
        else do;
          currsize + (buffIn-1);
          fmtLength = (buffIn-1);
        end;
        /* write only the bytes we read, no padding */
        outfmt = cats("$char", fmtLength, ".");
        rcPut = fput(fileid, putc(rec, outfmt));
        rcWrite = fwrite(fileid);      
        if (currsize >= &maxSize.) then leave;
      end;
      chunksize = currsize;
      bytescumulative + chunksize;
      byterange = cat("bytes ",bytescumulative-chunksize,"-",bytescumulative-1,"/",originalsize);
      output;
      rc = fclose(fileid);
    end;
    rc = fclose(filein);
  run;
%mend;

/* Upload a single file segment as part of an upload session */
%macro uploadFileChunk(
 uploadURL=,
 chunkFile=,
 byteRange=
);

  filename hdrout temp;
  filename resp temp;

  filename _tosave "&chunkFile.";
  proc http url= "&uploadURL"
     method="PUT"
     in=_tosave
     out=resp
     oauth_bearer="&access_token"
     headerout=hdrout
     ;
     headers
       "Content-Range"="&byteRange."
       ;
   run;

   %put NOTE: Upload segment &byteRange., HTTP result - &SYS_PROCHTTP_STATUS_CODE. &SYS_PROCHTTP_STATUS_PHRASE.; 

   /* HTTP 200 if success, 201 if new file was created */
  %if (%sysfunc(substr(&SYS_PROCHTTP_STATUS_CODE.,1,1)) ne 2) %then
    %do;
      %put WARNING: File upload failed!;
      %if (%sysfunc(fexist(resp))) %then
        %do;
          data _null_;
            rc=jsonpp('resp','log');
          run;
        %end;

      %if (%sysfunc(fexist(hdrout))) %then
        %do;
          data _null_;
            infile hdrout;
            input;
            put _infile_;
          run;
        %end;
    %end;

  filename _tosave clear;
  filename hdrout clear;
  filename resp clear;

%mend;

/* 
   Use an UploadSession in the Microsoft Graph API to upload a file.   

   This can handle large files, greater than the 4MB limit used by     
   PUT to the :/content endpoint.                                       
   The Graph API doc says you need to split the file into chunks.       

   We do need to know the total file size in bytes before using the API, so
   this code includes a file-size check.

   It also uses a splitFile macro to create a collection of file segments
   for upload. These must be in multiples of 320K size according to the doc
   (except for the last segment, which is a remainder size).
   
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
  filename resp_us temp;
 
   /* Create an upload session to upload the file.                                                */
   /* If a file of the same name exists, we will REPLACE it.                                      */
   /* The API doc says this should be POST, but since we provide a body with conflict directives, */
   /* it seems we must use PUT.                                                                   */
   proc http url="&msgraphApiBase./me/drives/&driveId./items/&folderId.:/%sysfunc(urlencode(&sourceFilename.)):/createUploadSession"
     method="PUT"
     in='{ "item": {"@microsoft.graph.conflictBehavior": "replace" }, "deferCommit": false }'
     out=resp_us
     ct="application/json"
     oauth_bearer="&access_token";
   run;
    %put NOTE: Create Upload Session: HTTP result - &SYS_PROCHTTP_STATUS_CODE. &SYS_PROCHTTP_STATUS_PHRASE.; 

    %if (&SYS_PROCHTTP_STATUS_CODE. = 200) %then %do;
      libname resp_us JSON fileref=resp_us;   
      data _null_;
      set resp_us.root; 
       call symputx('uploadURL',uploadUrl);
      run;
         
      %let fileSize=%getFileSize(localfile=&sourcePath./&sourceFilename.);  
   
      %put NOTE: Uploading &sourcePath./&sourceFilename., file size of &fileSize bytes.;

      /* split the file into segments for upload */
      %splitFile(
       sourceFile=&sourcePath./&sourceFilename.,
       maxSize = 1310720, /* 327680 * 4, must be multiples of 320K per doc */
       metadataOut=work._fileSegments
       );

      /* upload each segment file in this upload session */
      data _null_;
        set work._fileSegments;
        call execute(catt('%nrstr(%uploadFileChunk(uploadURL = %superq(uploadURL),chunkFile=',chunkPath,',byteRange=',byteRange,'));'));
      run;
      proc delete data=work._fileSegments;
    %end;
     /* Failed to create Upload Session */
     %else %do;
      %put WARNING: Upload session not created!; 
      %if (%sysfunc(fexist(resp_us))) %then %do;
        data _null_; rc=jsonpp('resp_us','log'); run;
      %end;
     %end;
     filename resp_us clear;
     options &_opt;
 %mend;
