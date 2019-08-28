
function getSSOToken() {
    Office.context.auth.getAccessTokenAsync(function (result) {
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
    });
});