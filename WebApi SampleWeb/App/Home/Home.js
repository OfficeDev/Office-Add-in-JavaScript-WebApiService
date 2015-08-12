/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#send').click(sendFeedback);
        });
    };

    function sendFeedback() {

        // Disable the controls while sending data
        $('.disable-while-sending').prop('disabled', true);

        var dataToPassToService = {
            Feedback: $('#feedback').val(),
            Rating: $('#rating').val()
        };

        $.ajax({
            url: '../../api/SendFeedback',
            type: 'POST',
            data: JSON.stringify(dataToPassToService),
            contentType: 'application/json;charset=utf-8'
        }).done(function (data) {
            app.showNotification(data.Status, data.Message);
        }).fail(function (status) {
            app.showNotification('Error', 'Could not communicate with the server.');
        }).always(function () {
            $('.disable-while-sending').prop('disabled', false);
        });
    }
})();


// *********************************************************
//
// Office-Add-in-JavaScript-WebApiService, https://github.com/OfficeDev/Office-Add-in-JavaScript-WebApiService
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************