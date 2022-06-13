$(document).ready(function() {
    if($("#perito_select").length>0) {
        $('#perito_select').select2();
    }
    $('#perito_select').on('change', function() {
        let idPerito = $(this).val();
        var formData = new FormData();
        formData.append("idPerito",idPerito);
        $.ajax({
            type: "POST",
            url: location.origin+"/pages/relatorios.php?status=ajax&case=1",
            async: true,
            data: formData,
            cache: false,
            success: function (msg) {

            },
            error: function (error) {
                console.log(error);
            },
            contentType: false,
            processData: false,
            enctype: 'multipart/form-data',
            timeout: 60000
        }).done(function (dir){
            openWindowWithPost(location.origin+"/pages/relatorios.php?status=ajax&case=2", {
                dir: dir,
            });
        })
    })
});

function openWindowWithPost(url, data) {
    var form = document.createElement("form");
    form.target = "_blank";
    form.method = "POST";
    form.action = url;
    form.style.display = "none";

    for (var key in data) {
        var input = document.createElement("input");
        input.type = "hidden";
        input.name = key;
        input.value = data[key];
        form.appendChild(input);
    }

    document.body.appendChild(form);
    form.submit();
    document.body.removeChild(form);
}