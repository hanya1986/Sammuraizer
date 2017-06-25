/* 
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 *  See LICENSE in the source repository root for complete license information.
 */

// This sample uses an open source OAuth 2.0 library that is compatible with the Azure AD v2.0 endpoint.
// Microsoft does not provide fixes or direct support for this library.
// Refer to the library’s repository to file issues or for other support.
// For more information about auth libraries see: https://azure.microsoft.com/documentation/articles/active-directory-v2-libraries/
// Library repo: http://adodson.com/hello.js/

(function () {
    angular
        .module('app')
        .controller('MainController', ['$http', '$scope', '$log', 'GraphHelper', function ($http, $scope, $log, GraphHelper) {

            // View model properties
            $scope.displayName;
            $scope.emailAddress;
            $scope.emails = [];

            /////////////////////////////////////////
            // End of exposed properties and methods.

            $scope.initAuth = function() {

                // Check initial connection status.
                if (localStorage.auth) {
                    $scope.processAuth();
                } else {
                    let auth = hello('aad').getAuthResponse();
                    if (auth !== null) {
                        localStorage.auth = angular.toJson(auth);
                        $scope.processAuth();
                    }
                }
            };


            // Auth info is saved in localStorage by now, so set the default headers and user properties.
            $scope.processAuth = function() {
                let auth = angular.fromJson(localStorage.auth);

                // Check token expiry. If the token is valid for another 5 minutes, we'll use it.
                let expiration = new Date();
                expiration.setTime((auth.expires - 300) * 1000);
                if (expiration > new Date()) {

                    // let the authProvider access the access token
                    authToken = auth.access_token;

                    // This header has been added to identify our sample in the Microsoft Graph service. If extracting this code for your project please remove.
                    $http.defaults.headers.common.SampleID = 'angular-connect-rest-sample';

                    if (localStorage.getItem('user') === null) {

                        // Get the profile of the current user.
                        GraphHelper.me().then(function(user) {

                            // Save the user to localStorage.
                            localStorage.setItem('user', angular.toJson(user));

                            $scope.displayName = user.displayName;
                            $scope.emailAddress = user.mail || user.userPrincipalName;
                        });
                    } else {
                        let user = angular.fromJson(localStorage.user);

                        $scope.displayName = user.displayName;
                        $scope.emailAddress = user.mail || user.userPrincipalName;
                    }
                }
            };

            $scope.fetchMails = function() {
                let auth = angular.fromJson(localStorage.auth);
                let expiration = new Date();
                expiration.setTime((auth.expires - 300) * 1000);
                if (expiration > new Date()) {

                    if (expiration > new Date()) {
                        GraphHelper.fetchMails().then(function (response) {
                            $log.debug('HTTP request to the Microsoft Graph API returned successfully.', response);
                            $scope.emails = response.value;
                            $scope.$apply();
                        }, function(error){
                            $log.error('HTTP request to the Microsoft Graph API failed.');
                        });
                    }
                }
            };

            $scope.refresh = function() {
                $scope.fetchMails();
            };
            $scope.initAuth();

            $scope.isAuthenticated = function() {
                return localStorage.getItem('user') !== null;
            };

            if ($scope.isAuthenticated()){
                $scope.fetchMails();
            };

            $scope.login = function() {
                GraphHelper.login();
            };

            $scope.logout = function() {
                GraphHelper.logout();
            };

        }]);
})();
