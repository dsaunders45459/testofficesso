
import hello from 'hellojs'
import $ from 'jquery'

// var /*@type(String)*/ ssoAadInstance = null;
// var /*@type(String)*/ tenantId = null;
// var /*@type(String)*/ clientId = '7fe71067-eb19-44ba-88cb-5317982f85d9';
// var /*@type(String)*/ ssoRedirectUrl = "https://localhost:3000/";
// var /*@type(String)*/ extraParameters = null;

// var config = {
//     tenant: tenantId,
//     instance: ssoAadInstance,
//     clientId: clientId,
//     redirectUri: ssoRedirectUrl,
//     navigateToLoginRequestUrl: false,
//     loadFrameTimeout: 30000,
//     popUp: true
// };

// var authContext = new AuthenticationContext(config);
// if (authContext.isCallback(window.location.hash)) {
//     authContext.handleWindowCallback();
// }

// function getAdalAccessToken() {
//     var resourceId = "api://localhost:3000/7fe71067-eb19-44ba-88cb-5317982f85d9/.default";
//     authContext.acquireTokenPopup(resourceId, null, null,  function (errorDesc, token, error) {
//         debugger;
//     });
//     // authContext.acquireToken(resourceId, function (acquireTokenError, token) {
//     //     debugger;
//     // });
// }

function testAdalJs() {
    // authContext.callback = function (loginError, idToken) {
    //     getAdalAccessToken(authContext);
    // }

    // let user = authContext.getCachedUser();
    // if (user) {
    //     getAdalAccessToken(authContext);
    // } else {
    //     authContext.login();
    // }
}

const AD_URL = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/';

hello.init({
    windows: "00000000481710A4"
}, {redirect_uri: "https://p.sfx.ms/sa.html", scope: "api://localhost:3000/7fe71067-eb19-44ba-88cb-5317982f85d9/.default", display: "popup"});
async function testHelloJS() {
    try
    {
        let result = await hello('windows').login({prompt: "none", display: "none"});
        console.log(result.authResponse);
    }
    catch(ex)
    {
        debugger;
    }
}


function getSSOToken() {
    Office.context.auth.getAccessTokenAsync(function (result) {
        if (result.status === "succeeded") {
            // Use this token to call Web API
            var ssoToken = result.value;
            $('#ssoToken').val(result.value);
        } else {
            $('#ssoToken').val(JSON.stringify(result));
        }
    });
}

function forceConsent() {
    Office.context.auth.getAccessTokenAsync({forceConsent:true}, function (result) {
        if (result.status === "succeeded") {
            // Use this token to call Web API
            var ssoToken = result.value;
            $('#ssoToken').val(result.value);
        } else {
            if (result.error.code === 13003) {
                // SSO is not supported for domain user accounts, only
                // work or school (Office 365) or Microsoft Account IDs.
            } else {
                // Handle error
            }
        }
    });
}

function getGraphToken() {
    $.ajax({type: "GET", 
		url: "/auth",
        headers: {"Authorization": "Bearer " + $('#ssoToken').val()},
        cache: false
    }).then(function (response) {
        $("#graphToken").val(JSON.stringify(response));
    });
}

function makeGraphApiCall() {
    $.ajax({type: "GET", 
        url: "/auth/getuserdata",
        headers: {"access_token":  JSON.parse($("#graphToken").val()).access_token},
        cache: false
    }).then(function (response) {
        $("#graphApiCall").val(JSON.stringify(response));
    });
}
function claimsRequest() {
    var claimsStr = JSON.parse($("#graphToken").val()).claims;
    Office.context.auth.getAccessTokenAsync({authChallenge: claimsStr}, function (result) {
        debugger;
    });
}
Office.onReady(function(info) {
    $(document).ready(function() {
        $('#getToken').click(getSSOToken);
        $('#forceConsent').click(forceConsent);
        $('#getGraphToken').click(getGraphToken);
        $('#makeGraphApiCall').click(makeGraphApiCall);
        $('#claims').click(claimsRequest);
        $('#testHelloJs').click(testHelloJS);
        $('#testAdalJs').click(testAdalJs);
    });
});