How to use the results of the Link Checker.

The output of the Link Checker will be in MS Excel.  The contains multiple columns, each are described below:
1. LinkID - This is an ID given to the link during the Link Checker process, it is just a number and is not needed for analysis.
2. Link Name - This is the display text that appears in the document for the link.  To fix a broken link, search for this text in the document and then a fix can be made to the actual URL.
3. Link URL - This is the actual URL or web address that the Link Name is trying to navigate to.
4. Link Check - This column will show a text description of the Link Check Results, for more information, please see the description below.
5. Check Code - This is the actual Get Request code returned, this is not really used for correcting URLs.

Link Check Code Details:

1. Valid - This means that the Link URL is a valid web page.
2. Redirect URL: Update - This means that the owners of the web page that the Link URL is navigating to have removed that web page and redirected users to a different web page.  The user should verify that the web page they are redirected to is the proper page, and then update the URL in the document.
3. Unauthorized: Check Manually - This means that the Link Checker was not authorized to visit the destination of the Link URL.  You should manually verify that this is the correct URL.
4. Forbidden: Check Manually - This means that the destination of the Link URL is not accessible without expressed permission based on user or IP.  You should manually verify that this is the correct URL.
5. No Response: Check Manually - This means that the check of the Link URL did not complete, this could mean that the web page is temporarily unavailable or the GSA firewall will not allow the connection.
6. Requires POST,GET,PULL - This means that the Link URL is pointing not to a web page, bubt to a server or API that is not meant to be accessed via a web browser.  This needs to be checked manually using the appropriate tools.
7. Invalid: Fix Link - This is an invalid request and means that the Link URL is most likely a bad link.
8. Internal Document Link - This means that the link is to an internal chapter, section, footnote, or something of that nature, not an actual webpage.