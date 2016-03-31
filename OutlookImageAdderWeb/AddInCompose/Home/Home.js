/// <reference path="../App.js" />

(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            debug('Initialized');

            // Roaming Settings
            var roamingValue = Office.context.roamingSettings.get("testvalue");
            if (!roamingValue) {
                debug("Setting roaming value");
                Office.context.roamingSettings.set("testvalue", "Roaming Settings Worked");
                Office.context.roamingSettings.saveAsync();
                roamingValue = Office.context.roamingSettings.get("testvalue");
            }
            debug("Roaming Test value is " + roamingValue);

            // Cookies
            var cookieValue = getCookie("testvalue");
            if (!cookieValue) {
                debug("Setting cookie value");
                setCookie("testvalue", "Cookies Worked");
                cookieValue = getCookie("testvalue");
            }
            debug("Cookie Test value is " + cookieValue);


            $('#attach-image-button').click(function () {
                debug('Attach Clicked');
                var src = $('#source-image').val();
                var method = $('#method').val();
                switch (method) {
                    case 'attach':
                        attachImage(src);
                        break;
                    case 'inline':
                        insertImage(src);
                        break;
                    default:
                        debug('Unknown method: ' + method);
                        break;
                }
            });
        });
    };

    function setCookie(key, value) {
        var expires = new Date();
        expires.setTime(expires.getTime() + (1 * 24 * 60 * 60 * 1000));
        document.cookie = key + '=' + value + ';expires=' + expires.toUTCString();
    }

    function getCookie(key) {
        var keyValue = document.cookie.match('(^|;) ?' + key + '=([^;]*)(;|$)');
        return keyValue ? keyValue[2] : null;
    }


    function debug(message) {
        $('#results').append('<p>' + message + '</p>');
    }

    function attachImage(srcUri) {
        debug('Attaching: ' + srcUri);

        Office.context.mailbox.item.addFileAttachmentAsync(srcUri, srcUri, {}, function() {});
        
        debug('Done');
    };

    function insertImage(srcUri) {
        debug('Inserting: ' + srcUri);
        var item = Office.context.mailbox.item;

        // TODO GET ACTUAL BASE64 ENCODED IMAGE
        var base64Image = "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAkACQAAD/2wBDAAIBAQIBAQICAgICAgICAwUDAwMDAwYEBAMFBwYHBwcGBwcICQsJCAgKCAcHCg0KCgsMDAwMBwkODw0MDgsMDAz/2wBDAQICAgMDAwYDAwYMCAcIDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAz/wAARCABfAGcDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD+f+iiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAP/Z";

        item.body.getTypeAsync(
            function (result) {
                if (result.status == Office.AsyncResultStatus.Failed) {
                    debug(result.error.message);
                }
                else {
                    debug('Body type is: ' + result.value);
                    // Successfully got the type of item body.
                    // Set data of the appropriate type in body.
                    if (result.value == Office.MailboxEnums.BodyType.Html) {
                        // Body is of HTML type.
                        // Specify HTML in the coercionType parameter
                        // of setSelectedDataAsync.

                               // Office.context.mailbox.item.addFileAttachmentAsync(srcUri, srcUri, {}, function() {});

                        
                        /*
                        console.log('<img src="' + base64Image + '"/>');

                        item.body.setSelectedDataAsync(
                            '<img src="' + base64Image + '"/>',
                            {
                                coercionType: Office.CoercionType.Html
                            },
                            function (asyncResult) {
                                if (asyncResult.status == "failed") {
                                    debug('Error: ' + asyncResult.error.message);
                                }
                            }
                        );

                        */
                    }
                    else {
                        // Body is of text type. 
                        // Cannot insert image into text type emails
                        debug("Error: Format is text, cannot insert inline images");
                    }
                }
            });

        debug('Done');
    };

})();