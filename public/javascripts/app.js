
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

Office.onReady((info) => {
    $(document).ready(() => {
        $('#getToken').click(getSSOToken);
        $('#forceConsent').click(forceConsent);
    });
});