$(document).ready( function () {
    if($("#myTable").length>0) {
        $('#myTable').DataTable({
            "ajax": {
                "url": $('#myTable').attr("ajaxpath"),
                "type": "POST"
            },
            "language": {
                "url": "//cdn.datatables.net/plug-ins/9dcbecd42ad/i18n/Portuguese-Brasil.json"
            },
            "order": [[ 0, "asc" ]],
            columnDefs: [
                { orderable: false, targets: [2] },
                {className: "dt-center", "targets": "_all"}
            ],
            "serverSide": true,
            "responsive": true,
            "processing": true
        });
    }

    $("#myTable").on("click","#infoFuncionario",function(e) {
        let idFuncionario = $(this).attr("data-idFuncionario");
        getInfoContatoColab(idFuncionario);
              
    })

} );

function getInfoContatoColab(idFuncionario){
    var formData = new FormData();
    formData.append("idFuncionario",idFuncionario);
    $.ajax({
        type: "POST",
        url: location.origin+"/pages/funcionario.php?status=ajax&case=3",
        async: true,
        data: formData,
        cache: false,
        success: function (msg) {
            msg = JSON.parse(msg);
            if(msg.status == 200){        
                var textResponse = "<h3>O Cliente não possui nenhum contato!</h3>";
                var have = "";
                if(msg.telefone != undefined){
                    textResponse = "<h3><b>Telefone: </b>"+msg.telefone+"</h3>";
                    var have = textResponse;
                }
                if(msg.email != undefined)
                    textResponse = have+"<h3><b>Email: </b>"+msg.email+"</h3>";

                Swal.fire({
                    title: 'Informações do '+ msg.nome +'!',
                    html:textResponse,
                    icon: 'info',
                    showConfirmButton: false,
                    showCancelButton: true,
                    cancelButtonText: 'Fechar'
                })
            }
            else{
                Swal.fire({
                    title: 'Erro!',
                    text: msg.message,
                    icon: 'error',
                    confirmButtonText: 'Ok'
                })
            }
        },
        error: function (error) {
            console.log(error);
        },
        contentType: false,
        processData: false,
        enctype: 'multipart/form-data'
    })
}