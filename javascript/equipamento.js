$(document).ready(function() {

    $('#add-equipamento').on('click', function(e){
        $.ajax({
            type: "POST",
            url: location.origin+"/pages/equipamento.php?status=ajax&case=1",
            async: false,
            cache: false,
            success: function (msg) {
                msg = JSON.parse(msg);
                if(msg.status == 200){
                    let modalProcesso = $('#modalEquipamento');
                    modalProcesso.find('.modal-header').html(msg.header);
                    modalProcesso.find('.modal-body').html(msg.body);
                    modalProcesso.find('.modal-footer').html(msg.footer);
                    modalProcesso.modal('show');

                }else{
                    Swal.fire({
                        title: 'Erro!',
                        text: msg.message,
                        icon: 'error',
                        confirmButtonText: 'Ok'
                    })
                }
            },
            contentType: false,
            processData: false,
        }).done(function (e){
            var table = $('#dataTables').DataTable({
                columnDefs: [
                    {type: 'monthYear', targets: 1 },
                    {className: "dt-center", "targets": "_all"}
                ],
                "ajax": {
                    "url": $('#dataTables').attr("ajaxpath"),
                    "type": "POST"
                },
                "drawCallback": function(settings) {
                    $('#dataTables').css("width", "100%")
                    $("button[name='downloadImagem']").on("click", function(e){
                        idEquipamento = $(this).attr("idEquipamento");
                        var formData = new FormData();
                        formData.append("idEquipamento",idEquipamento);
                        openWindowWithPost(location.origin+"/pages/equipamento.php?status=ajax&case=3", {
                            idEquipamento: idEquipamento,
                        });
                    })
                },
                
            });
        })

    })
})

function openWindowWithPost(url, data) {
    var form = document.createElement("form");
    // form.target = "_blank";
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
