
Dim ws As Object

	ws = GetWorkSheet(Handle, Workbook, Worksheet)

ws.AutoFilterMode = False
ws.Rows(RowTitle &":"& RowTitle).AutoFilter
ws.AutoFilter.Sort.SortFields.Clear
ws.AutoFilter.Sort.SortFields.Add(Key:=ws.Range(RangeFilter), SortOn:=0, Order:=1, DataOption:=0) ' Os números 1 correspondem a xlSortOnValues, xlAscending e xlSortNormal, respectivamente

With ws.AutoFilter.Sort
    .Header = 1 ' xlYes
    .MatchCase = False
    .Orientation = 1 ' xlTopToBottom
    .SortMethod = 1 ' xlPinYin
    .Apply
End With

' copiar codigo abaixo e colar a ação completa '

<process name="__selection__MS Excel VBO v2" type="object">
  <subsheet subsheetid="3c9a34b4-8c7d-4a86-b002-9866b0150457" type="Normal" published="True">
    <name>Sort Column</name>
    <view>
      <camerax>-58</camerax>
      <cameray>21</cameray>
      <zoom version="2">1.25</zoom>
    </view>
  </subsheet>
  <stage stageid="432d45a1-710c-493d-a553-1a6930954c4e" name="Sort Column" type="SubSheetInfo">
    <subsheetid>3c9a34b4-8c7d-4a86-b002-9866b0150457</subsheetid>
    <loginhibit onsuccess="true" />
    <narrative>Create By: Wellington Fonseca
Last update: Wellington Fonseca</narrative>
    <display x="-195" y="-105" w="150" h="90" />
    <font family="Tahoma" size="10" style="Regular" color="000000" />
  </stage>
  <stage stageid="84212545-1ab3-4e2c-b087-c3eac4106cd4" name="End" type="End">
    <subsheetid>3c9a34b4-8c7d-4a86-b002-9866b0150457</subsheetid>
    <loginhibit onsuccess="true" />
    <display x="15" y="-15" />
    <font family="Tahoma" size="10" style="Regular" color="000000" />
  </stage>
  <stage stageid="e54aa554-eb94-4dc0-ae94-3e33d21b5b84" name="Handle" type="Data">
    <subsheetid>3c9a34b4-8c7d-4a86-b002-9866b0150457</subsheetid>
    <loginhibit onsuccess="true" />
    <display x="-195" y="-30" w="150" h="30" />
    <font family="Tahoma" size="10" style="Regular" color="000000" />
    <datatype>number</datatype>
    <initialvalue />
    <private />
    <alwaysinit />
  </stage>
  <stage stageid="8785922c-4950-4a11-b88b-7eab7e40a627" name="Workbook" type="Data">
    <subsheetid>3c9a34b4-8c7d-4a86-b002-9866b0150457</subsheetid>
    <loginhibit onsuccess="true" />
    <display x="-195" y="0" w="150" h="30" />
    <font family="Tahoma" size="10" style="Regular" color="000000" />
    <datatype>text</datatype>
    <initialvalue />
    <private />
    <alwaysinit />
  </stage>
  <stage stageid="259d1335-38dc-4c93-a4cb-308290b6b2fa" name="Start" type="Start">
    <subsheetid>3c9a34b4-8c7d-4a86-b002-9866b0150457</subsheetid>
    <loginhibit onsuccess="true" />
    <preconditions>
      <condition narrative="Preencher os Date Itens e ter uma planilha aberta" />
    </preconditions>
    <postconditions>
      <condition narrative="Filtrará o Excel" />
    </postconditions>
    <display x="15" y="-105" />
    <font family="Tahoma" size="10" style="Regular" color="000000" />
    <inputs>
      <input type="number" name="Handle" narrative="is Handle" stage="Handle" />
      <input type="text" name="Workbook" narrative="is workbook name" stage="Workbook" />
      <input type="text" name="Worksheet" narrative="is worksheet name" stage="Worksheet" />
      <input type="number" name="RowTitle" narrative="Row of Title (cabecalho) exemple: 3" stage="RowTitle" />
      <input type="text" name="RangeFilter" narrative="column to filter, Exemple: B3 " stage="RangeFilter" />
    </inputs>
    <onsuccess>347b4d8e-5e2d-4c5c-b1dd-6ce67acfd6b4</onsuccess>
  </stage>
  <stage stageid="347b4d8e-5e2d-4c5c-b1dd-6ce67acfd6b4" name="Code" type="Code">
    <subsheetid>3c9a34b4-8c7d-4a86-b002-9866b0150457</subsheetid>
    <loginhibit onsuccess="true" />
    <display x="15" y="-60" />
    <font family="Tahoma" size="10" style="Regular" color="000000" />
    <inputs>
      <input type="number" name="Handle" expr="[Handle]" />
      <input type="text" name="Workbook" expr="[Workbook]" />
      <input type="text" name="Worksheet" expr="[Worksheet]" />
      <input type="number" name="RowTitle" expr="[RowTitle]" />
      <input type="text" name="RangeFilter" expr="[RangeFilter]" />
    </inputs>
    <onsuccess>84212545-1ab3-4e2c-b087-c3eac4106cd4</onsuccess>
    <code><![CDATA[
Dim ws As Object

	ws = GetWorkSheet(Handle, Workbook, Worksheet)

ws.AutoFilterMode = False
ws.Rows(RowTitle &":"& RowTitle).AutoFilter
ws.AutoFilter.Sort.SortFields.Clear
ws.AutoFilter.Sort.SortFields.Add(Key:=ws.Range(RangeFilter), SortOn:=0, Order:=1, DataOption:=0) ' Os números 1 correspondem a xlSortOnValues, xlAscending e xlSortNormal, respectivamente

With ws.AutoFilter.Sort
    .Header = 1 ' xlYes
    .MatchCase = False
    .Orientation = 1 ' xlTopToBottom
    .SortMethod = 1 ' xlPinYin
    .Apply
End With



]]></code>
  </stage>
  <stage stageid="1a6d1017-65a9-4b59-9d3b-b669cd51a82a" name="Worksheet" type="Data">
    <subsheetid>3c9a34b4-8c7d-4a86-b002-9866b0150457</subsheetid>
    <loginhibit onsuccess="true" />
    <display x="-195" y="30" w="150" h="30" />
    <font family="Tahoma" size="10" style="Regular" color="000000" />
    <datatype>text</datatype>
    <initialvalue />
    <private />
    <alwaysinit />
  </stage>
  <stage stageid="c73f6e24-94fd-491b-8a8a-8ac7346a77c8" name="RangeFilter" type="Data">
    <subsheetid>3c9a34b4-8c7d-4a86-b002-9866b0150457</subsheetid>
    <loginhibit />
    <display x="-195" y="75" w="150" h="30" />
    <datatype>text</datatype>
    <initialvalue />
    <private />
    <alwaysinit />
  </stage>
  <stage stageid="af0454bf-1826-4e0f-95ef-8a19ae2052ca" name="RowTitle" type="Data">
    <subsheetid>3c9a34b4-8c7d-4a86-b002-9866b0150457</subsheetid>
    <loginhibit />
    <display x="-195" y="105" w="150" h="30" />
    <datatype>number</datatype>
    <initialvalue />
    <private />
    <alwaysinit />
  </stage>
</process>
