var bar = new ProgressBar.Line('#container', {easing: 'easeInOut'});

$(document).ready(function() {
    $(document).on("click", "#salvaAlteracoes", function(e){
        var formData = new FormData();
        var tudoOk = true;
        formData.append("email",$("#email").val());
        formData.append("tel",$("#tel").val());
        
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
        if(tudoOk){
            Swal.fire({
                title: 'Por Favor, insira sua senha!',
                icon: 'info',
                input: 'password',
                inputPlaceholder: 'Insira sua senha',
                inputAttributes: {
                  maxlength: 10,
                  autocapitalize: 'off',
                  autocorrect: 'off'
                },
                showCancelButton: true,
                confirmButtonText: 'Salvar Alterações',
                showLoaderOnConfirm: true,
                preConfirm: (login) => {
                }
              }).then((result) => {
                if(result['isConfirmed']){
                    formData.append("pass",result['value']);
                    $.ajax({
                        type: "POST",
                        url: location.origin+"/pages/perfil.php?status=ajax&case=1",
                        async: true,
                        data: formData,
                        cache: false,
                        success: function (data) {
                            var data = JSON.parse(data);
                            if(data.status == 200){
                                if(data.passwordValid == true){
                                    alert("Dados alterados com sucesso. Atualizando a página para processar alterações...","","",3000);
                                    setTimeout(function(){ window.location.reload(); }, 3000);
                                }else{
                                    Swal.fire({
                                        title: 'Senha inválida!',
                                        icon: 'error',
                                        confirmButtonText: 'Ok'
                                    })
                                }
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
                }
            })
        }

    });

    $(document).on("click","#my-img",function(e) {
        $("input[id='img']").click();
    })
    $(document).on("change","#img",function(e) {
        var file = document.querySelector('input[type=file]').files[0];
        if(file != undefined){
            var formData = new FormData();
            formData.append('files', file);	
            Swal.fire({
                html: "Salvando Imagem. Por favor, aguarde...",
                showLoaderOnConfirm: true,
                allowOutsideClick: false,
            });
            Swal.showLoading();

            $.ajax({
                type: "POST",
                url: location.origin+"/pages/perfil.php?status=ajax&case=3",
                async: true,
                data: formData,
                cache: false,
                success: function (data) {
                    data = JSON.parse(data);
                    if(data.status == 200){
                        $('#imgSession').attr('src', data.imgSession);
                        previewFile(file,$("#my-img"));
                        alert(data.message);
                        Swal.close();
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
        }

    })
});

function previewFile(file,preview) {
    var reader  = new FileReader();
    reader.onloadend = function () {
      preview.attr("src",reader.result);
    }
  
    if (file) {
      reader.readAsDataURL(file);
    } else {
        preview.attr("src",'');
    }
  }