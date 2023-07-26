#Include "Protheus.ch"
#Include "TopConn.ch" 

/* 
 __________________________________________________________________________
| ///////////////////// TESTE CONSULTAS EXCEL /////////////////////////    |
|                                                                          |
|       Teste feito para consultas no Excel para a empresa Piffer.         |
| Func:ConsExl                                                             |
| Autor: Eduardo Alves                                                     |
|__________________________________________________________________________|
*/

#DEFINE ENTER CHR(13)+CHR(10)

USER FUNCTION RELEXCEL()

//Local oFWMsExcel
Local oFWMSExc
Local oExcel
//Local cArquivo    :=GETTEMPPATH() + 'RelExcel.xml'
Local aProdutos   := {}
Local aPedidos    := {}
Local nX          := 0
Local cArquivo  := "C:\Temp\teste_Edu.xml"

oFWMsExc := FWMsExcel():New()

//oFWMsExc:AddWorkSheet("Produtos")

  //FWMsExcelEx():AddTable( < cWorkSheet >, < cTable >)

Local cQry := ''

cQry := "SELECT A1_COD AS COD, A1_NOME AS NOME "
cQry += ", CASE WHEN A1_TIPO = 'R' THEN 'REVENDEDOR' "
cQry += "   WHEN A1_TIPO = 'F' THEN 'CONS. FINAL' "
cQry += "   ELSE 'X' END AS TIPO_CLI "
cQry += "FROM " +RetSqlName("SA1")+" "
cQry += "WHERE D_E_L_E_T_ = '' "

cQry := ChangeQuery(cQry)

TCQuery cQry New Alias "TabSA1"

        oFWMSExc:AddworkSheet("SA1X") 

        oFWMSExc:AddTable("SA1X","TESTESA1")
    
        oFWMSExc:AddColumn("SA1X","TESTESA1","COD",1,1)
        oFWMSExc:AddColumn("SA1X","TESTESA1","NOME",1,1)
        oFWMSExc:AddColumn("SA1X","TESTESA1","TIPO_CLI",1,1) 

Do While !TabSA1->(Eof())

        oFWMsExc:AddRow("SA1X","TESTESA1",{TabSA1->(COD),TabSA1->(NOME),TabSA1->(TIPO_CLI)})

        TabSA1->(DbSkip())

EndDo
oFWMSExc:Activate()
    oFWMSExc:GetXMLFile(cArquivo)
     
    //Abrindo o excel e abrindo o arquivo xml
    oExcel := MsExcel():New()             //Abre uma nova conexão com Excel
    oExcel:WorkBooks:open(cArquivo)       //Abre uma planilha
    oExcel:SetVisible(.T.)                //Visualiza a planilha
    oExcel:Destroy()                      //Encerra o processo do gerenciador de tarefas

Return
