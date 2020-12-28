

$(document).ready(function() {
    $("#getToken").click(() => {
        let iframe = document.createElement("iframe");
        iframe.src = "http://testdanssoapp2:3000";
        iframe.width = 800;
        iframe.height = 500;
        document.body.appendChild(iframe);
    });
    $("#forceConsent").click(async () => {
        let val = await document.hasStorageAccess();
        document.cookie = "test=1";
        $("#ssoToken").val(val.toString() + "-" + document.cookie);
    });
});
