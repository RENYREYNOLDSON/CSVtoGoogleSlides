# CSVtoGoogleSlides

**This code takes a CSV file and a Google Slide Template, it then appends the data to the template in a meaningful way.**\
*Developed for Advertize.net for integration with AWS lambda.*\
*Contact: reynoldson2002@gmail.com*\

## AWS REST API Endpoints
# Post Request
API Endpoint
~~~
https://pq2xyqsgla.execute-api.eu-north-1.amazonaws.com/testing
~~~
The passed data should be raw JSON containing {Google Access Token, Template File ID, Output File Name, CSV Data}
Example JSON Data
~~~
  
~~~
API Call Example
~~~
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
~~~
