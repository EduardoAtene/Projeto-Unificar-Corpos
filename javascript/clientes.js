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
                { orderable: false, targets: [3] },
                {className: "dt-center", "targets": "_all"}
            ],
            "serverSide": true,
            "responsive": true,
            "processing": true,
            "drawCallback": function(settings) {                         
                $("button[name='relatorioPeritoAnalise']").on('click', function(e){
                    let idClientePerito = $(this).attr("data-idClientePerito");

                    Swal.fire({
                        html: "Gerando Relatório. Por favor, aguarde...",
                        showLoaderOnConfirm: true,
                        allowOutsideClick: false,
                    })
                    Swal.showLoading();

                    var formData = new FormData();
                    formData.append("idClientePerito",idClientePerito);
                    
                    $.ajax({
                        type: "POST",
                        url: location.origin+"/pages/relatorios.php?status=ajax&case=6",
                        data: formData,
                        async: true,
                        cache: false,
                        success: function (data) {
                            var data = JSON.parse(data);
                            if(data.status == 200){
                                var $a = $("<a>");
                                $a.attr("href",data.file);
                                $("body").append($a);
                                $a.attr("download",data.fileName+".xlsx");
                                $a[0].click();
                                $a.remove();
            
                            }else{
                                alert("Ocorreu um problema ao emitir o relatório. Por Favor, tende mais tarde!");
                            }
                            Swal.close();
                        },
                        error: function (error) {
                            console.log(error);
                        },
                        contentType: false,
                        processData: false,
                        enctype: 'multipart/form-data',
                    });
                })
                    

            },
        });
    }

    $("#myTable").on("click","#infoPerito",function(e) {
        let idClientePerito = $(this).attr("data-idCliente");
        getInfoContatoColab(idClientePerito);
              
    })

} );

function getInfoContatoColab(idCliente){
    var formData = new FormData();
    formData.append("idCliente",idCliente);
    $.ajax({
        type: "POST",
        url: location.origin+"/pages/clientes.php?status=ajax&case=3",
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
                    title: 'Informações '+ msg.nome +'!',
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