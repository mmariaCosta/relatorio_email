#Include "Protheus.ch"
#Include "FwMVCDef.ch"

// Cadastro de Emails
User Function EMAIL() 

    Local cFunBkp   := FunName()
    Local oBrowse   := Nil

    SetFunName("EMAIL")
    ChkFile("SZB")
    
    oBrowse := FwMBrowse():New()
    oBrowse:setAlias('SZB')
    oBrowse:setMenuDef("EMAIL")
    oBrowse:setDescription('Cadastro de Emails')
    oBrowse:Activate()

    SetFunName(cFunBkp)

Return

Static Function ModelDef()
    Local oStruSZB  := FwFormStruct(1,'SZB')
    Local oModel

    oModel := MPFormModel():New("EMAILMODEL")
    oModel:AddFields('SZBMASTER', /*cOwner*/, oStruSZB)
    oModel:SetPrimaryKey({'ZB_FILIAL', 'ZB_COD'})
    oModel:SetDescription('Manutenção de Emails')

Return oModel

Static Function ViewDef()

    Local oModel    := FWLoadModel('EMAIL')
    Local oStruSZB  := FWFormStruct(2,'SZB')
    Local oView

    oView := FWFormView():New()
    oView:SetModel(oModel)
    oView:AddField('VIEW_SZB', oStruSZB, 'SZBMASTER')
    oView:CreateHorizontalBox('CABEC', 100)
    oView:SetOwnerView('VIEW_SZB', 'CABEC')

Return oView

Static Function MenuDef()

    Local aRotina := {}

    ADD OPTION aRotina Title 'Incluir'      Action 'VIEWDEF.EMAIL' OPERATION 3 ACCESS 0
    ADD OPTION aRotina Title 'Alterar'      Action 'VIEWDEF.EMAIL' OPERATION 4 ACCESS 0
    ADD OPTION aRotina Title 'Visualizar'   Action 'VIEWDEF.EMAIL' OPERATION 2 ACCESS 0
    ADD OPTION aRotina Title 'Excluir'      Action 'VIEWDEF.EMAIL' OPERATION 5 ACCESS 0
    ADD OPTION aRotina Title 'Imprimir'     Action 'VIEWDEF.EMAIL' OPERATION 8 ACCESS 0
    ADD OPTION aRotina Title 'Copiar'       Action 'VIEWDEF.EMAIL' OPERATION 9 ACCESS 0

Return aRotina
