<process name="__selection__SAP - FBL1N Parameters" type="object">
  <subsheet subsheetid="ce82f1cc-923f-47dd-a725-b881a5c36981" type="Normal" published="True">
    <name>Input All Parameters</name>
    <view>
      <camerax>0</camerax>
      <cameray>-12</cameray>
      <zoom version="2">1.25</zoom>
    </view>
  </subsheet>
  <stage stageid="6a508a37-2c92-4b65-9157-4c0ac88b879e" name="Input All Parameters" type="SubSheetInfo">
    <subsheetid>ce82f1cc-923f-47dd-a725-b881a5c36981</subsheetid>
    <narrative>Faz o input de todos os parametros menos de seleções multiplas.

Created by: Wellington
Last Changed: Wellington</narrative>
    <display x="-195" y="-105" w="150" h="90" />
  </stage>
  <stage stageid="32b47ffa-a6b0-43b8-b30f-a0c5200b35b6" name="Start" type="Start">
    <subsheetid>ce82f1cc-923f-47dd-a725-b881a5c36981</subsheetid>
    <loginhibit />
    <preconditions>
      <condition narrative="Necessário estar na janela principal" />
    </preconditions>
    <postconditions>
      <condition narrative="Faz o input de todos os parametros" />
    </postconditions>
    <display x="15" y="-105" />
    <inputs>
      <input type="text" name="Conta de Fornecedor" narrative="Input do elemento" stage="Coll.Conta_de_fornecedor" />
      <input type="text" name="Conta de fornecedor_Ate" narrative="Input do elemento" stage="Coll.Conta_de_fornecedor_Ate" />
      <input type="text" name="Empresa" narrative="Input do elemento" stage="Coll.Empresa" />
      <input type="text" name="Empresa_Ate" narrative="Input do elemento" stage="Coll.Empresa_Ate" />
      <input type="text" name="ID ajud pesq" narrative="Input do elemento" stage="Coll.ID_ajud_pesq" />
      <input type="text" name="Cad pesq" narrative="Input do elemento" stage="Coll.Cad_pesq" />
      <input type="flag" name="Partidas em aberto" narrative="Input do elemento" stage="Coll.Partidas_em_aberto" />
      <input type="text" name="Aberto a data fixada_Partidas em aberto" narrative="Input do elemento" stage="Coll.Aberto_a_data fixada_Partidas_em_aberto" />
      <input type="text" name="Data de compensação" narrative="Input do elemento" stage="Coll.Data_de_compensação" />
      <input type="flag" name="Partidas compensadas" narrative="Input do elemento" stage="Coll.Partidas_compensadas" />
      <input type="text" name="Data de compensação_Ate" narrative="Input do elemento" stage="Coll.Data_de_compensação_Ate" />
      <input type="text" name="Aberto à data fixada_Partidas Compensadas" narrative="Input do elemento" stage="Coll.Aberto_à_data fixada_Partidas_Compensadas" />
      <input type="flag" name="Todas as partidas" narrative="Input do elemento" stage="Coll.Todas_as_partidas" />
      <input type="text" name="Data de lançamento" narrative="Input do elemento" stage="Coll.Data_de_lançamento" />
      <input type="text" name="Data de lançamento_Ate" narrative="Input do elemento" stage="Coll.Data_de_lançamento_Ate" />
      <input type="flag" name="Partidas normais" narrative="Input do elemento" stage="Coll.Partidas_normais" />
      <input type="flag" name="Operações do Razão Especial" narrative="Input do elemento" stage="Coll.Operações_do_Razão_Especial" />
      <input type="flag" name="Partida-memo" narrative="Input do elemento" stage="Coll.Partida-memo" />
      <input type="flag" name="Partidas pré-editadas" narrative="Input do elemento" stage="Coll.Partidas_pré-editadas" />
      <input type="flag" name="Partida em débito" narrative="Input do elemento" stage="Coll.Partida_em_débito" />
      <input type="text" name="Layout" narrative="Input do elemento" stage="Coll.Layout" />
      <input type="text" name="Número máximo de partidas" narrative="Input do elemento" stage="Coll.Número_máximo_de_partidas" />
    </inputs>
    <onsuccess>74c555e0-8c60-4b69-a9b3-c34fd90eed03</onsuccess>
  </stage>
  <stage stageid="8a665c95-331c-408b-902b-bda3196e0c40" name="End" type="End">
    <subsheetid>ce82f1cc-923f-47dd-a725-b881a5c36981</subsheetid>
    <loginhibit />
    <display x="15" y="-15" />
  </stage>
  <stage stageid="74c555e0-8c60-4b69-a9b3-c34fd90eed03" name="Input all Itens" type="Code">
    <subsheetid>ce82f1cc-923f-47dd-a725-b881a5c36981</subsheetid>
    <loginhibit />
    <display x="15" y="-60" />
    <inputs>
      <input type="collection" name="Coll" expr="[Coll]" />
    </inputs>
    <onsuccess>8a665c95-331c-408b-902b-bda3196e0c40</onsuccess>
    <code><![CDATA[Dim session As Object
Dim Application As Object
Dim Connection As Object
Dim SapGuiAuto As Object

SapGuiAuto  = GetObject("SAPGUI")
Application = SapGuiAuto.GetScriptingEngine
Connection = Application.Children(0)
session = Connection.Children(0)

if Coll.Rows(0)("Partidas_em_aberto") = True
session.findById("wnd[0]/usr/radX_OPSEL").select
end if
if Coll.Rows(0)("Partidas_compensadas") = True
session.findById("wnd[0]/usr/radX_CLSEL").select
end if
if Coll.Rows(0)("Todas_as_partidas") = True
session.findById("wnd[0]/usr/radX_AISEL").select
end if

session.findById("wnd[0]/usr/chkX_NORM").selected = Coll.Rows(0)("Partidas_normais")
session.findById("wnd[0]/usr/chkX_SHBV").selected = Coll.Rows(0)("Operações_do_Razão_Especial")
session.findById("wnd[0]/usr/chkX_MERK").selected = Coll.Rows(0)("Partida-memo")
session.findById("wnd[0]/usr/chkX_PARK").selected = Coll.Rows(0)("Partidas_pré-editadas")
session.findById("wnd[0]/usr/chkX_APAR").selected = Coll.Rows(0)("Partida_em_débito")
session.findById("wnd[0]/usr/ctxtKD_LIFNR-LOW").text = Coll.Rows(0)("Conta_de_fornecedor")
session.findById("wnd[0]/usr/ctxtKD_LIFNR-HIGH").text = Coll.Rows(0)("Conta_de_fornecedor_Ate")
session.findById("wnd[0]/usr/ctxtKD_BUKRS-LOW").text = Coll.Rows(0)("Empresa")
session.findById("wnd[0]/usr/ctxtKD_BUKRS-HIGH").text = Coll.Rows(0)("Empresa_Ate")
session.findById("wnd[0]/usr/ctxtKD_INDEX-HOTKEY").text = Coll.Rows(0)("ID_ajud_pesq")
session.findById("wnd[0]/usr/ctxtKD_INDEX-STRING").text = Coll.Rows(0)("Cad_pesq")
session.findById("wnd[0]/usr/ctxtPA_STIDA").text = Coll.Rows(0)("Aberto_a_data fixada_Partidas_em_aberto")
session.findById("wnd[0]/usr/ctxtSO_AUGDT-LOW").text = Coll.Rows(0)("Data_de_compensação")
session.findById("wnd[0]/usr/ctxtSO_AUGDT-HIGH").text = Coll.Rows(0)("Data_de_compensação_Ate")
session.findById("wnd[0]/usr/ctxtPA_STID2").text = Coll.Rows(0)("Aberto_à_data fixada_Partidas_Compensadas")
session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = Coll.Rows(0)("Data_de_lançamento")
session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = Coll.Rows(0)("Data_de_lançamento_Ate")
session.findById("wnd[0]/usr/ctxtPA_VARI").text = Coll.Rows(0)("Layout")
session.findById("wnd[0]/usr/txtPA_NMAX").text = Coll.Rows(0)("Número_máximo_de_partidas")
]]></code>
  </stage>
  <stage stageid="7b73ce13-705e-496b-8b1c-aa24f9a0bba3" name="Coll" type="Collection">
    <subsheetid>ce82f1cc-923f-47dd-a725-b881a5c36981</subsheetid>
    <loginhibit />
    <display x="-195" y="-15" w="120" h="30" />
    <datatype>collection</datatype>
    <private />
    <alwaysinit />
    <collectioninfo>
      <field name="Conta_de_fornecedor" type="text" />
      <field name="Empresa" type="text" />
      <field name="ID_ajud_pesq" type="text" />
      <field name="Cad_pesq" type="text" />
      <field name="Aberto_à_data fixada_Partidas_Compensadas" type="text" />
      <field name="Data_de_compensação" type="text" />
      <field name="Conta_de_fornecedor_Ate" type="text" />
      <field name="Empresa_Ate" type="text" />
      <field name="Data_de_compensação_Ate" type="text" />
      <field name="Aberto_a_data fixada_Partidas_em_aberto" type="text" />
      <field name="Data_de_lançamento" type="text" />
      <field name="Layout" type="text" />
      <field name="Número_máximo_de_partidas" type="text" />
      <field name="Data_de_lançamento_Ate" type="text" />
      <field name="Partidas_normais" type="flag" />
      <field name="Operações_do_Razão_Especial" type="flag" />
      <field name="Partida-memo" type="flag" />
      <field name="Partidas_pré-editadas" type="flag" />
      <field name="Partida_em_débito" type="flag" />
      <field name="Partidas_em_aberto" type="flag" />
      <field name="Partidas_compensadas" type="flag" />
      <field name="Todas_as_partidas" type="flag" />
    </collectioninfo>
    <initialvalue>
      <row>
        <field name="Conta_de_fornecedor" type="text" value="" />
        <field name="Empresa" type="text" value="" />
        <field name="ID_ajud_pesq" type="text" value="" />
        <field name="Cad_pesq" type="text" value="" />
        <field name="Aberto_à_data fixada_Partidas_Compensadas" type="text" value="" />
        <field name="Data_de_compensação" type="text" value="" />
        <field name="Conta_de_fornecedor_Ate" type="text" value="" />
        <field name="Empresa_Ate" type="text" value="" />
        <field name="Data_de_compensação_Ate" type="text" value="" />
        <field name="Aberto_a_data fixada_Partidas_em_aberto" type="text" value="" />
        <field name="Data_de_lançamento" type="text" value="" />
        <field name="Layout" type="text" value="" />
        <field name="Número_máximo_de_partidas" type="text" value="" />
        <field name="Data_de_lançamento_Ate" type="text" value="" />
        <field name="Partidas_normais" type="flag" value="" />
        <field name="Operações_do_Razão_Especial" type="flag" value="" />
        <field name="Partida-memo" type="flag" value="" />
        <field name="Partidas_pré-editadas" type="flag" value="" />
        <field name="Partida_em_débito" type="flag" value="" />
        <field name="Partidas_em_aberto" type="flag" value="" />
        <field name="Partidas_compensadas" type="flag" value="" />
        <field name="Todas_as_partidas" type="flag" value="" />
      </row>
    </initialvalue>
  </stage>
  <stage stageid="2e45e3c3-615d-4eda-8897-f6f4eed721c9" name="Inputs" type="Block">
    <subsheetid>ce82f1cc-923f-47dd-a725-b881a5c36981</subsheetid>
    <loginhibit />
    <display x="-270" y="-45" w="150" h="60" />
    <font family="Segoe UI" size="10" style="Regular" color="7FB2E5" />
  </stage>
</process>
