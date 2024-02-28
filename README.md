# CSVtoGoogleSlides
### This repository documents the development of Synchronous and Asynchronous serverless functions developed using AWS Lambda.
*Developed for Advertize.net for integration with AWS lambda.*\
*Contact: reynoldson2002@gmail.com*



## Accessing the Google Slides API
**This code takes a CSV file and a Google Slide Template, it then appends the data to the template in a meaningful way.**\
index.html has example code to access the API from a web app (Must be run on http://127.0.0.1:5000).
There is a dedicated Google Cloud Project for this application as we need to read and write to Google Slide files. Currently it is not published so your gmail account will need to be on the testing user's list to access. To add an account, email reynoldson2002@gmail.com and it will be added ASAP. On logging in to google through the app portal an access token will be created.

**Google Client ID:** '341236422016-lcv56vfvqpcsdoqq6ahjpa6ghrpn57rc.apps.googleusercontent.com'\
**Google App ID:** '341236422016'\
**Required Scopes:** 'https://www.googleapis.com/auth/drive.metadata.readonly https://www.googleapis.com/auth/presentations https://www.googleapis.com/auth/drive'

The **Google Picker API** is a great way to select slide files within a web app. It provides a window with the familiar Google Design and allows us to return the ID of the required template file. Documentation of this can be found here https://developers.google.com/drive/picker/guides/overview. To use this, the page URL must be added on the Google Cloud Project Console.

## AWS REST API's
## Synchronous Invocation
We can send the data to an AWS REST API which passes it through directly to AWS Lambda. However, the timeout for a response is 30 seconds and this function can take several minutes. An AWS Lambda function can take up to 15 minutes, so it can still perform the function on AWS but no response is provided to the client. It's also possible to directly call the lambda function from it's resource URL for this.
#### Example Using this API:
1. Client sends POST request containing specified data in the body to the REST API (CSVtoSlidesAPI).
2. This triggers a lambda function (CSVtoSlides) which attempts to run the CSVtoSlides function.
3. No response is given to the client.
   
![Syncrhonous REST API Architecture](https://miro.medium.com/v2/resize:fit:1400/format:webp/1*ZN6LVw0r9Po3asAqgu2oLg.png)
*Image: AWS Synchronous REST API Schematic*
### Post Request
API Endpoint
~~~
https://pq2xyqsgla.execute-api.eu-north-1.amazonaws.com/testing
~~~
The passed data should be raw JSON containing 
1. Google Access Token
2. Template File ID
3. Output File Name
4. CSV Data
   
Example JSON Data
```js
{
"accessToken":"ya29.a0AfB_byCpXNY9ZOryFUqdK_Ru3kVMNtyQFKK6djp5R1QbMqd0Qo8a5eN17D1K9NJajNZtj4_KoZX9pMT7Rs169",
"fileID":"1JTViOt75gUXKLHf-sOYe_CPsvThV7J0AaKjWa9LwlIw",
"outputName":"OUTPUT File",
"CSVdata":{"data":[
  ["Slide Type","Table Type","Title","Table Name","Target Calls","Data","","","","","","","","","","","",""],
  ["Skip","","","Title Slide","","","","","","","","","","","","","",""],
  ["Single","Ripple","All Accounts","Combined","3528","10/08/2023","$214,463","98@$2,188","0.50","1,778@$121","$138,182","71@$1,946","0.57"],
  ["Single","Ripple","All Google","","","10/08/2023","$138,182","","71@$1,946","0.57","1,778@$78","(,25,46,)","$722,119","","480@$1,504","0.77"],                          
  ["Single","Ripple","Goog/Criminal","","","10/08/2023","$65,433","13@$5,033","17@$3,849","0.34","615@$106","(,13,4,)","$278,217"]
}
```
API Call Example
```js
  fetch("https://pq2xyqsgla.execute-api.eu-north-1.amazonaws.com/testing/", {
      method: "POST",
      body: JSON.stringify({
          accessToken,
          fileID,
          outputName,
          CSVdata
  }),
      headers: {
          "Content-type": "application/json",
          "Accept":"*/*"
  }
  })
```
Sometimes the API will take several minutes to process and get data from Google. Therefore no response is given, look at asynchronous invocation to get a response.

## Asynchronous Invocation
In order to get a response from the REST API on a long run time function we must call the AWS lambda function asynchronously. To achieve this it is recommended to save the POST request into AWS's DynamoDB first, then have the lambda function triggered by this new item being added. The client then polls the API with a GET request containing the requests unique ID, until they get a valid response.
#### Example Using this API:
1. Client sends POST request containing specified data in the body to the REST API (CSVtoSlidesAPI-ASYNCH). The request-id (key) is returned to the client. NOTE: The CSV data must be encoded to store in DynamoDB so the 2D array is joined using ¶ and ¦ as delimiters.
2. A mapping template maps the recieved request data into columns for DynamoDB, and creates a new item (containing the request data) in the linked DynamoDB table (CSVtoSlides_Table). The item uses the request-id as the primary key and has a 'Complete' field to track if the lambda function has completed (initially this is False).
3. When the new request is added to DynamoDB it triggers an attatched lambda function (CSVtoSlides-ASYNCH). This runs the CSVtoSlides function and updates the table with a link to the google slides file or an error message when applicable.
4. The client starts polling the API using a GET request containing the request-id. This GET request runs a lambda function (CSVtoSlides_Get_Completed) which finds the item with request-id as the key and returns the 'Completed' field. Once the client recieves a response not equal to "False", we know that the main lambda function has finished and we can stop polling (or if timed out).
   
![Asyncrhonous REST API Architecture](https://d2908q01vomqb2.cloudfront.net/fc074d501302eb2b93e2554793fcaf50b3bf7291/2021/05/06/Figure-1.jpg)
*Image: AWS Asynchronous REST API Schematic (we bypass using the S3 service)*
### Post Request with Polling
API Endpoint
~~~
https://pq2xyqsgla.execute-api.eu-north-1.amazonaws.com/testing
~~~
The passed data should be raw JSON containing 
1. Google Access Token
2. Template File ID
3. Output File Name
4. CSV Data (Encoded so that we can store in DynamoDB)
   
Example JSON Data
```js
{
"accessToken":"ya29.a0AfB_byCpXNY9ZOryFUqdK_Ru3kVMNtyQFKK6djp5R1QbMqd0Qo8a5eN17D1K9NJajNZtj4_KoZX9pMT7Rs169",
"fileID":"1JTViOt75gUXKLHf-sOYe_CPsvThV7J0AaKjWa9LwlIw",
"outputName":"OUTPUT File",
"CSVdata":"Slide Type¶Table Type¶Title¶Table Name¶Target Calls¶Data¶¶¶¶¶¶¶¶¶¶¦Skip¶¶¶Title Slide¶¶¶¶¶¶¶¶¶¶¶¶¦Single¶Ripple¶All Accounts¶Combined¶6421¶10/12/2023¶13,125@1.99¶$788¶6,502@1.70¶$3,821¶6,623@2.39¶$2,766¶¶¶¶¦Single¶Ripple¶All Google¶¶¶11/12/2033¶$3,821¶44¶6,556@1.70¶$14,532¶245¶22,532@1.54¶0@$0¶0@$0¶6,502@$1¶22,432@$1¦"
}
```
API Call Example
```js
   CSVdata=CSVdata.map(innerArray => innerArray.join('¶')).join('¦');
   fetch("https://u5ydjdjy0j.execute-api.eu-north-1.amazonaws.com/test", {
       method: "POST",
       body: JSON.stringify({
           accessToken,
           fileID,
           outputName,
           CSVdata
   }),
       headers: {
           "Content-type": "application/json",
           "Accept":"*/*"
   }
   })
   .then((response) => response.json())
   .then((json) => {
     key = json["context"]["request-id"];
     const startTime = new Date().getTime();
     pollingFunction(key,startTime);
   });
```
Polling Function
```js
   //POLL UNTIL A NON FALSE RESULT IS REACHED (the slide code has finished processing, return a link)
   function pollingFunction(key,startTime){
     const currentTime = new Date().getTime();
     const elapsedTime = currentTime - startTime;
     document.getElementById("responseText").innerText = "POST Request Sent - Waiting on Response... Polling Time Elapsed: "+elapsedTime/1000+"s"
     if (elapsedTime >= 300000) { // 5 minutes = 300,000 milliseconds
       console.log("Polling stopped after 5 minutes.");
       return; // Stop polling
     };
     fetch("https://u5ydjdjy0j.execute-api.eu-north-1.amazonaws.com/test?key="+key, {
                   method: "GET",
                   headers: {
                       "Content-type": "application/json",
                       "Accept":"*/*",
               }
               })
               .then((response) => response.json())
               .then((json) => {
                 if (json["body"] !== '"False"') {
                   // If response is not 'False', do something with the response or stop polling
                   if (json["body"].includes("https://")){
                     document.getElementById("responseText").innerText = "SUCCESSFUL RESPONSE:\nLink: "+json["body"]+"\nPolling Time Elapsed: "+elapsedTime/1000+"s"
                   }
                   else{
                     document.getElementById("responseText").innerText = "FAILED RESPONSE:\n"+json["body"]+"\nPolling Time Elapsed: "+elapsedTime/1000+"s"
                   };
                   
                 } else {
                   // If response is 'False', continue polling after a delay of 10 seconds
                   setTimeout(() => {
                       pollingFunction(key, startTime);
                   }, 10000); // 10000 milliseconds = 10 seconds
                 }
               })
               .catch((error) => {
                 console.error('Error:', error);
                 // Retry polling after a delay of 10 seconds in case of error
                 setTimeout(() => {
                     pollingFunction(key, startTime);
                 }, 10000); // 10000 milliseconds = 10 seconds
             });;
   };


```
Look at index.html to understand the format of the responses.

## Additional Implementation Documentation


