$(document).ready( function () {

    var table = $('#dataTables').DataTable({
        "ajax": {
            "url": location.origin+"/pages/usuarios.php?status=ajax&case=1",
            "type": "POST"
          },
          "language": {
            "url": "//cdn.datatables.net/plug-ins/9dcbecd42ad/i18n/Portuguese-Brasil.json"
          },
          columnDefs: [
            {"className": "dt-center", "targets": "_all"}
        ],
    });

    $('#newFuncionario').on('click',function () {
        $("#auxStade").val("1");
        let modalCadastro = $('#modal');
        $.ajax({
            method: "POST",
            async: true,
            url: location.origin+"/pages/usuarios.php?status=ajax&case=2",
            data: {"tipoAcesso":"funcionario","isedit":false}
        }).done(function (msg) {
            msg = JSON.parse(msg);
            if(msg.status == 200){
                modalCadastro.find('.modal-header').html(msg.header);
                modalCadastro.find('.modal-body').html(msg.body);
                modalCadastro.find('.modal-footer').html(msg.footer);
                modalCadastro.modal('show');
            }
        });
    })

    $('#newPersonal').on('click',function () {
        $("#auxStade").val("1");
        let modalCadastro = $('#modal');
        $.ajax({
            method: "POST",
            async: true,
            url: location.origin+"/pages/usuarios.php?status=ajax&case=2",
            data: {"tipoAcesso":"personal","isedit":false}
        }).done(function (msg) {
            msg = JSON.parse(msg);
            if(msg.status == 200){
                modalCadastro.find('.modal-header').html(msg.header);
                modalCadastro.find('.modal-body').html(msg.body);
                modalCadastro.find('.modal-footer').html(msg.footer);
                modalCadastro.modal('show');
            }
        });
    })

    $('#newCliente').on('click',function () {
        $("#auxStade").val("1");
        let modalCadastro = $('#modal');    
        $.ajax({
            method: "POST",
            async: true,
            url: location.origin+"/pages/usuarios.php?status=ajax&case=2",
            data: {"tipoAcesso":"cliente","isedit":false}
        }).done(function (msg) {
            msg = JSON.parse(msg);
            if(msg.status == 200){
                modalCadastro.find('.modal-header').html(msg.header);
                modalCadastro.find('.modal-body').html(msg.body);
                modalCadastro.find('.modal-footer').html(msg.footer);
                modalCadastro.modal('show');

                    
                $('#cpf').on('blur',function () {
                    var cpf = $("#cpf").val();
                    var formData = new FormData();
                    formData.append('cpf', cpf);	
                    if(!(cpf==''||cpf==null || cpf.length < 13)){
                        $.ajax({
                            type: "POST",
                            async: true,
                            url: location.origin+"/pages/usuarios.php?status=ajax&case=3",
                            data: formData,                            
                            cache: false,
                            contentType: false,
                            processData: false,
                            enctype: 'multipart/form-data',
                        }).done(function (msg) {
                            msg = JSON.parse(msg);
                            if(msg.status == 200){
                                $("#nome").val(msg.nome);
                            }else{
                                $("#nome").val();
                            }
                        });
                    }else{
                        if($("#nome").val()==''||$("#nome").val()==null){
                            $("#nome").val(msg.nome);
                        }else
                            $("#nome").val(msg.nome);
                    }
                })
            }
        });
    })



    $(document).on("click", "button[data-dismiss='salvaForm']", function(e){
        var formData = new FormData();
        tudoOk = true;
        if($('#img').length > 0){
            if($('#img')[0].files.length > 0){
                formData.append("files", $('#img')[0].files[0]);
            }
        }
        if($("#usuario").length>0){
            if($("#usuario").val()==''||$("#usuario").val()==null){
                $("#usuario").attr('class','form-control is-invalid')
                tudoOk = false;
            }
            else{
                formData.append("usuario", $("#usuario").val());
                $("#usuario").attr('class','form-control is-valid')
            }
        }
            
        if($("#senha").length>0){
            if($("#senha").val()==''||$("#senha").val()==null){
                $("#senha").attr('class','form-control is-invalid')
                tudoOk = false;
            }
            else{
                formData.append("senha", $("#senha").val());
                $("#senha").attr('class','form-control is-valid')
            }
        }

        if($("#nome").length>0){
            if($("#nome").val()==''||$("#nome").val()==null){
                $("#nome").attr('class','form-control is-invalid')
                tudoOk = false;
            }
            else{
                formData.append("nomeCompleto", $("#nome").val());
                $("#nome").attr('class','form-control is-valid')
            }
        }

        if($("#cpf").length>0){
            if($("#cpf").val()==''||$("#cpf").val()==null || $("#cpf").val().length < 13){
                $("#cpf").attr('class','form-control is-invalid')
                tudoOk = false;
            }
            else{
                formData.append("cpf", $("#cpf").val());
                $("#cpf").attr('class','form-control is-valid')
            }
        }

        if($("#email").length>0){
            if($("#email").val()==''||$("#email").val()==null){
                $("#email").attr('class','form-control is-invalid')
                tudoOk = false;
            }
            else{
                formData.append("email", $("#email").val());
                $("#email").attr('class','form-control is-valid')
            }
        }

        if($("#tel").length>0){
            if($("#tel").val()==''||$("#tel").val()==null || $("#tel").val().length < 14){
                $("#tel").attr('class','form-control is-invalid')
                tudoOk = false;
            }
            else{
                formData.append("tel", $("#tel").val());
                $("#tel").attr('class','form-control is-valid')
            }
        }
        if($("#cargo").length>0){
            if($("#cargo").val()==''||$("#cargo").val()==null){
                $("#cargo").attr('class','form-control is-invalid')
            }
            else{
                formData.append("cargo", $("#cargo").val());
                $("#cargo").attr('class','form-control is-valid')
            }
        }


        if($(this).attr('data-infoid')!=undefined)
            formData.append("infoId",$(this).attr('data-infoid'));
        formData.append("tipoUsuario",$(this).attr("data-typeForm"));

        if(tudoOk){
            $.ajax({
                type: "POST",
                url: location.origin+"/pages/perfil.php?status=ajax&case=1",
                async: true,
                data: formData,
                cache: false,
                success: function (data) {
                    if(data==-1){ //usuario ja existe
                        Swal.fire({
                            title: 'Erro ao realizar o cadastro!',
                            text: 'Já existe um cadastro com esse usuário',
                            icon: 'error',
                            confirmButtonText: 'Ok'
                        })
                    }
                    else{
                        let modalCadastro = $('#modal');
                        let textAux;
                        if($("#auxStade").val() == "1"){
                            textAux = 'Cadastro realizado com sucesso';
                        }else{
                            textAux = 'Edição realizada com sucesso';
                        }
                        Swal.fire({
                            title: textAux,
                            icon: 'success',
                            confirmButtonText: 'Ok'
                        }).then((result) => {
                            /* Read more about isConfirmed, isDenied below */
                            if (result.isConfirmed) {
                                modalCadastro.modal('hide');
                                table.ajax.reload();
                                setTimeout(function () { location.reload(); }, 2000);
                            }
                        })
                    }
                    
                },
                error: function (error) {
                    console.log(error);
                    let modalCadastro = $('#modal');
                    Swal.fire({
                        title: 'Erro ao realizar o cadastro',
                        text: 'Clique para continuar',
                        icon: 'error',
                        confirmButtonText: 'Ok'
                    }).then((result) => {
                        /* Read more about isConfirmed, isDenied below */
                        if (result.isConfirmed) {
                            modalCadastro.modal('hide');
                            table.ajax.reload();
                        }
                    })
                },
                contentType: false,
                processData: false,
                enctype: 'multipart/form-data',
                timeout: 60000
            });
        }

        return true;
    });

    $(document).on("click", "button[name='edit-user']", function(e){
        $("#auxStade").val("2");
        var formData = new FormData();
        let tableHeaderCells = ($(this).closest('table').find('thead').find('tr').find('th'));
        let indexName, indexCPF
        let name,cpf
        tableHeaderCells.each(function(index){
            if($( this ).text().trim() == "Nome"){
                indexName = index
            }
            else if($( this ).text().trim() == "CPF"){
                indexCPF = index
            }
        })
        tableLine = $(this).closest('tr').children('td');
        name = (tableLine.eq(indexName).text().trim());
        cpf = (tableLine.eq(indexCPF).text().trim());
        formData.append("nomeCompleto", name);
        formData.append("cpf",cpf);
        formData.append("isedit",true);
        let modalCadastro = $('#modal');
        $.ajax({
            type: "POST",
            url: location.origin+"/pages/usuarios.php?status=ajax&case=2",
            async: true,
            data: formData,
            cache: false,
            success: function (msg) {
                msg = JSON.parse(msg);
                if(msg.status == 200){
                    modalCadastro.find('.modal-header').html(msg.header);
                    modalCadastro.find('.modal-body').html(msg.body);
                    modalCadastro.find('.modal-footer').html(msg.footer);
                    if($("#plano_select").length>0)
                        $('#plano_select').select2({
                            width: '100%' // need to override the changed default
                        });
                    $('#regiao_select').select2({
                            width: '100%' // need to override the changed default
                        });
                    $('#nome').keyup(function(){
                        $('#nome').val($("#nome").val().toUpperCase());
                    });
                    modalCadastro.modal('show');
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
            enctype: 'multipart/form-data',
            timeout: 60000
        });
        return true;
    });
    
} );