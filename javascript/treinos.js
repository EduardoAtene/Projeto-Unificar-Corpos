$(document).ready(function() {

    $('#add-modalidade').on('click', function(e){
        var formData = new FormData();
        formData.append("hash",-1);
        formData.append("listHashChecked",$("#valuesCheckElement").val());
        $.ajax({
            type: "POST",
            url: location.origin+"/pages/modalidades.php?status=ajax&case=1",
            async: false,
            data: formData,
            cache: false,
            success: function (msg) {
                msg = JSON.parse(msg);
                if(msg.status == 200){
                    let modalProcesso = $('#modalModalidade');
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
            debugger
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
                        idModalidade = $(this).attr("idModalidade");
                        var formData = new FormData();
                        formData.append("idModalidade",idModalidade);
                        openWindowWithPost(location.origin+"/pages/modalidades.php?status=ajax&case=3", {
                            idModalidade: idModalidade,
                        });
                    })
                },
                
            });
        })

    })

    $('#add-turma').on('click', function(e){
        e.preventDefault();
        $.ajax({
            type: "POST",
            url: location.origin+"/pages/modalidades.php?status=ajax&case=4",
            async: false,
            cache: false,
            success: function (msg) {
                msg = JSON.parse(msg);
                if(msg.status == 200){
                    let modalProcesso = $('#modalTurma');
                    modalProcesso.find('.modal-header').html(msg.header);
                    modalProcesso.find('.modal-body').html(msg.body);
                    modalProcesso.find('.modal-footer').html(msg.footer);
                    modalProcesso.modal('show');
                    
                    $('#buttonSaveTurma').on('click',function(e){

                        let idModalidade = $("#modalidade").val();
                        let funcionarioResp = $("#funcionarioResp").val();
                        let submit = true;
                        if(idModalidade == -1){
                            $("#modalidade").removeClass("is-valid").addClass("is-invalid");
                            e.preventDefault();
                            submit = false;
                        }else{
                            $("#idModalidade").val(funcionarioResp);
                            $("#modalidade").removeClass("is-invalid").addClass("is-valid");
                        }
                        if(funcionarioResp == -1){
                            $("#funcionarioResp").removeClass("is-valid").addClass("is-invalid");
                            e.preventDefault();
                            submit = false;
                        }else{
                            $("#idFuncionarioResp").val(funcionarioResp);
                            $("#funcionarioResp").removeClass("is-invalid").addClass("is-valid");
                        }

                        if(submit)
                            e.submit();
                    });

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
        })

    })

    $(document).on("click", 'a[input-send="buttonCard"]', function(){
        debugger
        let nameModalidade = $(this).attr("input-name");
        let idTurma = $(this).attr("input-id");
        let desinscrever = $(this).attr("input-desinscrever");
        var textFirst, textConfirmButton, textSecond;
        if(desinscrever != "false"){
            textFirst = 'Deseja deinscrever do '+nameModalidade+'?';
            textConfirmButton = 'Deinscrever';
            textSecond = 'Inscrição Cancelada';
        }else{
            textFirst = 'Deseja se Inscrever no '+nameModalidade+'?';
            textConfirmButton = 'Inscrever';
            textSecond = 'Inscrito';
        }
        var formData = new FormData();
        formData.append("idTurma",idTurma);
        formData.append("desinscrever",desinscrever);

        Swal.fire({
            title: textFirst,
            showCancelButton: true,
            confirmButtonText: textConfirmButton,
            icon: 'question'
          }).then((result) => {
            if (result.isConfirmed) {
                $.ajax({
                    type: "POST",
                    url: location.origin+"/pages/modalidades.php?status=ajax&case=5",
                    async: false,
                    cache: false,
                    data: formData,
                    success: function (msg) {
                        msg = JSON.parse(msg);
                        if(msg.status == 200){
                            Swal.fire(textSecond+'!', '', 'success').then((result)=>{
                                location.reload();
                            })
                        }else{
                            Swal.fire({
                                title: 'Erro!',
                                text: 'Ocorreu um erro ao se Inscrever. Tente mais tarde',
                                icon: 'error',
                                confirmButtonText: 'Ok'
                            }).then((result)=>{
                                location.reload();
                            });
                        }
                    },
                    error: function (error) {
                        console.log(error);
                    },
                    contentType: false,
                    processData: false,
                })

            }
          })
    });

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
