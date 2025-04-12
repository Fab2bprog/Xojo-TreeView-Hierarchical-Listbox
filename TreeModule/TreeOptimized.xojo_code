#tag Class
Protected Class TreeOptimized
Inherits DesktopListBox
	#tag Event
		Sub CellAction(row As Integer, column As Integer)
		  Var CellTarget_LevelNomencl as int64
		  Var CellTarget_KeyNomencl as int64
		  Var CellIsChecked as Int64 =0
		  Var SqlQuery as string
		  
		  // [Update-XojoCode-NextLines]
		  Var ps      As SQLitePreparedStatement 
		  
		  CellTarget_LevelNomencl =me.CellTextAt(row,1).CDbl
		  CellTarget_KeyNomencl   =me.CellTextAt(row,2).CDbl
		  
		  if me.CellCheckBoxStateAt(row,column)=DesktopCheckBox.VisualStates.Checked then
		    CellIsChecked=1
		  else
		    CellIsChecked=0
		  end if
		  
		  // [Update-XojoCode-NextLines]
		  SqlQuery  = " SELECT * FROM "+SrcDataObject.TableName+" WHERE  NCLLEVEL= ? AND KEY_ID= ? "
		  
		  ps = SrcDataObject.DBaseID.Prepare(SqlQuery)
		  ps.BindType(0, SQLitePreparedStatement.SQLITE_INT64)
		  ps.BindType(1, SQLitePreparedStatement.SQLITE_INT64)
		  ps.Bind(0, CellTarget_LevelNomencl)
		  ps.Bind(1,  CellTarget_KeyNomencl)
		  
		  SrcDataObject.Fields_Init
		  SrcDataObject.DBaseRS=ps.SelectSQL
		  
		  Call SrcDataObject.Record_ReadFirst
		  SrcDataObject.Field_IsSelected=CellIsChecked
		  Call SrcDataObject.Record_Update
		  
		  
		End Sub
	#tag EndEvent

	#tag Event
		Function ConstructContextualMenu(base As DesktopMenuItem, x As Integer, y As Integer) As Boolean
		  if me.SelectedRowCount >0 Then
		    
		    if me.CellTextAt(me.SelectedRowIndex,1).ToInt64>0 then
		      Ite_AddBrother      = New DesktopMenuItem
		      Ite_AddBrother.text  = "Add an element of the same level"
		      Ite_AddBrother.Name  = "Ite_AddBrother"
		      Ite_AddBrother.Enabled=True
		      base.AddMenu Ite_AddBrother
		    end if
		    
		    
		    Ite_AddChild       =  New DesktopMenuItem
		    Ite_AddChild.text  = "Add a lower level element"
		    Ite_AddChild.Name  = "Ite_AddChild"
		    Ite_AddChild.Enabled=True
		    base.AddMenu Ite_AddChild
		    
		    
		    if me.CellTextAt(me.SelectedRowIndex,1).ToInt64>0 then
		      Ite_RenameNode        =  New DesktopMenuItem
		      Ite_RenameNode.text   = "Rename this element "
		      Ite_RenameNode.Name   = "Ite_RenameNode"
		      Ite_RenameNode.Enabled=True
		      base.AddMenu Ite_RenameNode
		      
		      Ite_DelNode       =  New DesktopMenuItem
		      Ite_DelNode.text  = "Delete the element and its contents"
		      Ite_DelNode.Name  = "Ite_DelNode"
		      Ite_DelNode.Enabled=True
		      base.AddMenu Ite_DelNode
		      
		      Ite_SelectEmoji      =  New DesktopMenuItem
		      Ite_SelectEmoji.text  = "Select an Emoji"
		      Ite_SelectEmoji.Name  = "Ite_SelectEmoji"
		      Ite_SelectEmoji.Enabled=True
		      base.AddMenu Ite_SelectEmoji
		      
		    end if
		    
		  end if
		  
		  
		  return true
		  
		End Function
	#tag EndEvent

	#tag Event
		Sub RowExpanded(row As Integer)
		  me.Tree_Build(me.CellTextAt(row,1).ToInt64+1,me.CellTextAt(row,2).ToInt64)
		  
		  
		End Sub
	#tag EndEvent


	#tag MenuHandler
		Function Ite_AddBrother() As Boolean Handles Ite_AddBrother.Action
		  Var TextWithoutEmoji as string
		  TextWithoutEmoji = str(me.CellTextAt(me.SelectedRowIndex,5))
		  
		  
		  Win_InputBox.Parametre("Creation of an element of the same nomenclature level as '"+TextWithoutEmoji+"'","")
		  Win_InputBox.ShowModal
		   
		  If WinInput_IsOk then
		    me.Tree_Node_Add(WinInput_InputValue,true)
		  end if
		  
		  return true
		  
		  
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function Ite_AddChild() As Boolean Handles Ite_AddChild.Action
		  Var TextWithoutEmoji as string
		  TextWithoutEmoji = str(me.CellTextAt(me.SelectedRowIndex,5))
		  
		  Win_InputBox.Parametre("Creation of a lower-level element than '" +TextWithoutEmoji+"'","")
		  Win_InputBox.ShowModal
		  
		  If  WinInput_IsOk then
		    me.Tree_Node_Add(WinInput_InputValue,false)
		  end if
		  
		  return true
		  
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function Ite_DelNode() As Boolean Handles Ite_DelNode.Action
		  Var NodeLevel as Int64
		  Var NodeKey as Int64
		  
		  Var d As New MessageDialog                   ' declare the MessageDialog object
		  Var b As MessageDialogButton                 ' for handling the result
		  
		  if me.SelectedRowCount =0 Then
		    MessageBox "To use this function you must first select an item in the list"
		    exit Function
		  end if
		  
		  d.IconType = MessageDialog.IconTypes.Caution ' display warning icon
		  d.ActionButton.Caption = "Delete"
		  d.CancelButton.Visible = True                ' show the Cancel button
		  d.Message = "Delete the selected item?"
		  d.Explanation = "This will result in the deletion of any items they may contain."
		  
		  b = d.ShowModal                          ' display the dialog
		  Select Case b                                ' determine which button was pressed.
		  Case d.ActionButton
		    NodeKey = me.CellTextAt(me.SelectedRowIndex,2).ToInt64
		    NodeLevel = me.CellTextAt(me.SelectedRowIndex,1).ToInt64
		    
		    me.Tree_Hierarchy_Delete(NodeKey,NodeLevel)
		    
		    me.RowExpandedAt(me.SelectedRowIndex)=False
		    me.RemoveRowAt(me.SelectedRowIndex)
		    
		  Case d.CancelButton
		    exit Function
		  End Select
		  
		  
		  
		  
		  
		  
		  Return True
		  
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function Ite_Exit() As Boolean Handles Ite_Exit.Action
		  me.Close
		  Return True
		  
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function Ite_RenameNode() As Boolean Handles Ite_RenameNode.Action
		  Var TextWithoutEmoji as string
		  TextWithoutEmoji = str(me.CellTextAt(me.SelectedRowIndex,5))
		  
		  Win_InputBox.Parametre("Choose a new name for this element",TextWithoutEmoji)
		  Win_InputBox.ShowModal
		  
		  If WinInput_IsOk then
		    me.Tree_Node_Rename(WinInput_InputValue)
		  end if
		  
		  return true
		  
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function Ite_SelectEmoji() As Boolean Handles Ite_SelectEmoji.Action
		  Var TextCharEmoji as string
		  TextCharEmoji = str(me.CellTextAt(me.SelectedRowIndex,4))
		  
		  Win_InputBoxEmoji.Parametre("Choose a new name for this element",TextCharEmoji)
		  Win_InputBoxEmoji.ShowModal
		  
		  If WinInput_IsOk then
		    me.Tree_Node_SelEmoji(WinInput_Emoji)
		  end if
		  
		  return true
		  
		End Function
	#tag EndMenuHandler


	#tag Method, Flags = &h0
		Sub Tree_Build(NodeLevel as Int64, NodeKey as Int64)
		  Var SqlQuery as string
		  
		  // [Update-XojoCode-NextLines]
		  Var ps      As SQLitePreparedStatement 
		  
		  SrcDataObject.SqlQuerySource=""
		  
		  Select case NodeLevel
		  Case 0
		    me.RemoveAllRows
		    me.AddExpandableRow str("All")
		    me.CellTextAt(me.LastAddedRowIndex,1) = str(0 )
		    me.CellTextAt(me.LastAddedRowIndex,2) = str(0 )
		    me.CellTextAt(me.LastAddedRowIndex,3) = str(0 )
		    me.CellCheckBoxStateAt(0,0)=DesktopCheckBox.VisualStates.Checked
		    exit sub
		  Case 1
		    
		    // [Update-XojoCode-NextLines]
		    SqlQuery = " SELECT * FROM "+SrcDataObject.TableName+" WHERE NCLLEVEL = ? ORDER BY NAME  "
		    
		    ps = SrcDataObject.DBaseID.Prepare(SqlQuery)
		    ps.BindType(0, SQLitePreparedStatement.SQLITE_INT64)
		    ps.Bind(0,1)
		    
		  Case is >=2
		    
		    // [Update-XojoCode-NextLines]
		    SqlQuery  = " SELECT * FROM "+SrcDataObject.TableName+" WHERE  NCLPARENT = ? AND NCLLEVEL= ? ORDER BY NCLLEVEL,NCLPARENT,NAME "
		    
		    ps = SrcDataObject.DBaseID.Prepare(SqlQuery)
		    ps.BindType(0, SQLitePreparedStatement.SQLITE_INT64)
		    ps.BindType(1, SQLitePreparedStatement.SQLITE_INT64)
		    ps.Bind(0,NodeKey)
		    ps.Bind(1,NodeLevel)
		    
		  Case Else
		    exit sub
		  end select
		  
		  SrcDataObject.Fields_Init
		  SrcDataObject.DBaseRS=ps.SelectSQL
		  
		  if SrcDataObject.DBaseRS.RowCount=0 then
		    return
		  end if
		  
		  Call  SrcDataObject.Record_ReadFirst
		  
		  while not SrcDataObject.DBaseRS.AfterLastRow
		    
		    if me.Manage_Emoji then
		      me.AddExpandableRow SrcDataObject.Field_Emoji+SrcDataObject.Field_Name
		    else
		      me.AddExpandableRow SrcDataObject.Field_Name
		    end if
		    
		    If me.Manage_CheckBox then
		      me.CellTypeAt(me.LastAddedRowIndex,0)=DesktopListbox.CellTypes.CheckBox
		      if SrcDataObject.Field_IsSelected=0 then
		        me.CellCheckBoxStateAt(me.LastAddedRowIndex,0) =DesktopCheckBox.VisualStates.Unchecked
		      else
		        me.CellCheckBoxStateAt(me.LastAddedRowIndex,0) =DesktopCheckBox.VisualStates.Checked
		      end if
		    end if
		    
		    me.CellTextAt(me.LastAddedRowIndex,1) = SrcDataObject.Field_NclLevel.ToString
		    me.CellTextAt(me.LastAddedRowIndex,2) = SrcDataObject.KeyTableValue.ToString
		    me.CellTextAt(me.LastAddedRowIndex,3) = SrcDataObject.Field_NclParent.ToString
		    me.CellTextAt(me.LastAddedRowIndex,4) = SrcDataObject.Field_Emoji
		    me.CellTextAt(me.LastAddedRowIndex,5) = SrcDataObject.Field_Name
		    
		    
		    if 0=NodeLevel and 0=SrcDataObject.KeyTableValue then
		      me.CellCheckBoxValueAt(me.LastAddedRowIndex,0)=True
		    end if
		    
		    
		    Call  SrcDataObject.Record_ReadNext
		    
		  wend
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function Tree_Exist_Child(Parent as Int64, ParentLevel as Int64) As Boolean
		  Var SqlQuery as string
		  Var rs       as RowSet
		  
		  // [Update-XojoCode-NextLines]
		  Var ps      As SQLitePreparedStatement 
		  
		  // [Update-XojoCode-NextLines]
		  SqlQuery= "SELECT COUNT(*) AS NBRELEMENT FROM "+SrcDataObject.TableName+"  WHERE NCLPARENT = ? AND NCLLEVEL = ?  "
		  ps = SrcDataObject.DBaseID.Prepare(SqlQuery)
		  ps.BindType(0, SQLitePreparedStatement.SQLITE_INT64)
		  ps.BindType(1, SQLitePreparedStatement.SQLITE_INT64)
		  ps.Bind(0,Parent)
		  ps.Bind(1,ParentLevel+1)
		  
		  
		  Try
		    rs = ps.SelectSQL
		  Catch error As DatabaseException
		    MessageBox("Bad Sql resquest: " + error.Message)
		    return false
		  End Try
		  
		  Try
		    rs.MoveToFirstRow
		  Catch error As DatabaseException
		    MessageBox("DB Error: " + error.Message)
		    return false
		  End Try
		  
		  if rs.Column("NBRELEMENT").Value >0 then
		    return true
		  end if
		  
		  return false
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function Tree_FirstChildKey_Find(Parent as Int64, ParentLevel as Int64) As Int64
		  Var SqlQuery as string
		  Var rs       as RowSet
		  
		  // [Update-XojoCode-NextLines]
		  Var ps      As SQLitePreparedStatement 
		  
		  // [Update-XojoCode-NextLines]
		  SqlQuery= "SELECT KEY_ID FROM "+SrcDataObject.TableName+" WHERE NCLPARENT = ? AND NCLLEVEL = ? LIMIT 1 "
		  ps = SrcDataObject.DBaseID.Prepare(SqlQuery)
		  ps.BindType(0, SQLitePreparedStatement.SQLITE_INT64)
		  ps.BindType(1, SQLitePreparedStatement.SQLITE_INT64)
		  ps.Bind(0,Parent)
		  ps.Bind(1,ParentLevel+1)
		  
		  
		  
		  Try
		    rs = ps.SelectSQL
		  Catch error As DatabaseException
		    MessageBox("Bad Sql resquest: " + error.Message)
		    return -1
		  End Try
		  
		  Try
		    rs.MoveToFirstRow
		  Catch error As DatabaseException
		    MessageBox("DB Error: " + error.Message)
		    return -1
		  End Try
		  
		  
		  return rs.Column("KEY_ID").Int64Value
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Tree_Hierarchy_Delete(NodeKey as int64, NodeLevel as int64)
		  Var NodeCible as Int64
		  
		  
		  do until not Tree_Exist_Child(NodeKey,NodeLevel)
		    NodeCible = Tree_FirstChildKey_Find(NodeKey,NodeLevel)
		    Tree_Hierarchy_Delete(NodeCible,NodeLevel+1)
		  loop
		  
		  Call Tree_Node_Delete(NodeKey,NodeLevel)
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Tree_Node_Add(NewNodeName as String, AddBrother as Boolean = true)
		  
		  Var NclParent   as Int64
		  Var NclNodeLevel   as Int64
		  Var NodeKey   as Int64
		  
		  
		  if AddBrother  then
		    NclParent = me.CellTextAt(me.SelectedRowIndex,3).ToInt64
		    NclNodeLevel = me.CellTextAt(me.SelectedRowIndex,1).ToInt64
		  else
		    NclParent = me.CellTextAt(me.SelectedRowIndex,2).ToInt64
		    NclNodeLevel = me.CellTextAt(me.SelectedRowIndex,1).ToInt64+1
		  end if
		  
		  if Tree_Node_Compliant("ADD",NewNodeName,NclParent,NclNodeLevel)=false then
		    exit sub
		  end if
		  
		  SrcDataObject.Fields_Init()
		  SrcDataObject.Field_Name          = NewNodeName
		  SrcDataObject.Field_NclLevel        = NclNodeLevel
		  SrcDataObject.Field_NclParent     = NclParent
		  Call SrcDataObject.Record_Create
		  
		  NodeKey = SrcDataObject.KeyTableValue
		  
		  if AddBrother  then
		    me.AddExpandableRowAt(me.SelectedRowIndex,SrcDataObject.Field_Emoji+NewNodeName,NclNodeLevel)
		    
		    if me.Manage_CheckBox then 
		      me.CellTypeAt(me.SelectedRowIndex-1,0)=DesktopListbox.CellTypes.CheckBox
		      me.CellCheckBoxStateAt(me.SelectedRowIndex-1,0) =DesktopCheckBox.VisualStates.Unchecked
		    end if
		    
		    me.CellTextAt(me.SelectedRowIndex-1,1) = str(NclNodeLevel)
		    me.CellTextAt(me.SelectedRowIndex-1,2) = str(NodeKey)
		    me.CellTextAt(me.SelectedRowIndex-1,3) = str(NclParent)
		    me.CellTextAt(me.SelectedRowIndex-1,4) = str(SrcDataObject.Field_Emoji)
		    me.CellTextAt(me.SelectedRowIndex-1,5) = NewNodeName
		  else
		    if  me.RowExpandedAt(me.SelectedRowIndex) then
		      me.AddExpandableRowAt(me.SelectedRowIndex+1,SrcDataObject.Field_Emoji+NewNodeName,NclNodeLevel)
		      
		      if me.Manage_CheckBox then 
		        me.CellTypeAt(me.SelectedRowIndex+1,0)=DesktopListbox.CellTypes.CheckBox
		        me.CellCheckBoxStateAt(me.SelectedRowIndex+1,0) =DesktopCheckBox.VisualStates.Unchecked
		      end if
		      
		      me.CellTextAt(me.SelectedRowIndex+1,1) = str(NclNodeLevel)
		      me.CellTextAt(me.SelectedRowIndex+1,2) = str(NodeKey)
		      me.CellTextAt(me.SelectedRowIndex+1,3) = str(NclParent)
		      me.CellTextAt(me.SelectedRowIndex+1,4) = str(SrcDataObject.Field_Emoji)
		      me.CellTextAt(me.SelectedRowIndex+1,5) = NewNodeName
		    else
		      me.RowExpandedAt(me.SelectedRowIndex)=true
		    end if
		  end if
		  
		  
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1000
		Function Tree_Node_Compliant(TypeModif as string, ElementName as string, Parent as int64, NodeLevel as Int64) As Boolean
		  Var SqlQuery as string
		  Var rs       as RowSet
		  
		  // [Update-XojoCode-NextLines]
		  Var ps      As SQLitePreparedStatement 
		  
		  if ElementName.trim.Length=0 then
		    MessageBox "The name of this element is invalid"
		    return false
		  end if
		  
		  
		  Select Case TypeModif
		  Case "ADD"
		    
		    // [Update-XojoCode-NextLines]
		    SqlQuery= "SELECT  COUNT(*) AS NBRELEMENT FROM "+SrcDataObject.TableName+" WHERE UPPER(NAME) = UPPER(?)  AND NCLPARENT = ? AND NCLLEVEL = ? "
		    
		    ps = SrcDataObject.DBaseID.Prepare(SqlQuery)
		    ps.BindType(0, SQLitePreparedStatement.SQLITE_TEXT)
		    ps.BindType(1, SQLitePreparedStatement.SQLITE_INT64)
		    ps.BindType(2, SQLitePreparedStatement.SQLITE_INT64)
		    ps.Bind(0,ElementName)
		    ps.Bind(1,Parent)
		    ps.Bind(2,NodeLevel)
		    
		    
		  Case "UPDATE"
		    
		    // [Update-XojoCode-NextLines]
		    SqlQuery= "SELECT  COUNT(*) AS NBRELEMENT FROM "+SrcDataObject.TableName+" WHERE "+SrcDataObject.KeyTableName+_
		    " <> ? AND UPPER(NAME) = UPPER(?)  AND NCLPARENT = ? AND NCLLEVEL = ? "
		    
		    ps = SrcDataObject.DBaseID.Prepare(SqlQuery)
		    ps.BindType(0, SQLitePreparedStatement.SQLITE_INT64)
		    ps.BindType(1, SQLitePreparedStatement.SQLITE_TEXT)
		    ps.BindType(2, SQLitePreparedStatement.SQLITE_INT64)
		    ps.BindType(23, SQLitePreparedStatement.SQLITE_INT64)
		    ps.Bind(0,me.CellTextAt(me.SelectedRowIndex,2).ToInt64)
		    ps.Bind(1,ElementName)
		    ps.Bind(2,Parent)
		    ps.Bind(3,NodeLevel)
		    
		  Case Else
		    return false
		  end select
		  
		  Try
		    rs=ps.SelectSQL
		  Catch error As DatabaseException
		    MessageBox("Bad Sql resquest: " + error.Message)
		    return false
		  End Try
		  
		  Try
		    rs.MoveToFirstRow
		  Catch error As DatabaseException
		    MessageBox("DB Error: " + error.Message)
		    return false
		  End Try
		  
		  if rs.Column("NBRELEMENT").Value >0 then
		     MessageBox("An element already has an identical name in this location!")
		    return false
		  end if
		  
		  
		  return true
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function Tree_Node_Delete(NodeKey as Int64, NodeLevel as Int64) As Boolean
		  Var SqlQuery as string
		  
		  // [Update-XojoCode-NextLines]
		  Var ps      As SQLitePreparedStatement 
		  
		  // [Update-XojoCode-NextLines]
		  SqlQuery=" DELETE FROM "+SrcDataObject.TableName+" WHERE "+SrcDataObject.KeyTableName+" = ? AND NCLLEVEL = ? "
		  
		  ps = SrcDataObject.DBaseID.Prepare(SqlQuery)
		  ps.BindType(0, SQLitePreparedStatement.SQLITE_INT64)
		  ps.BindType(1, SQLitePreparedStatement.SQLITE_INT64)
		  ps.Bind(0,NodeKey)
		  ps.Bind(1,NodeLevel)
		  
		  Try
		    ps.ExecuteSQL
		  Catch error As DatabaseException
		    MessageBox("DB Error: " + error.Message)
		    return False
		  End Try
		  
		  
		  
		  return True
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Tree_Node_Rename(NewNodeName as String)
		  Var SqlQuery as string
		  
		  Var NodeKey as Int64
		  Var NodeLevel as Int64
		  Var Parent as Int64
		  
		  // [Update-XojoCode-NextLines]
		  Var ps As SQLitePreparedStatement 
		  
		  NodeLevel = me.CellTextAt(me.SelectedRowIndex,1).ToInt64
		  NodeKey = me.CellTextAt(me.SelectedRowIndex,2).ToInt64
		  Parent = me.CellTextAt(me.SelectedRowIndex,3).ToInt64
		  
		  if Tree_Node_Compliant("UPDATE",NewNodeName,Parent,NodeLevel) then
		    Win_InputBox.Close
		  else
		    Win_InputBox.Txt_NodeName.SetFocus
		    exit sub
		  end if
		  
		  // [Update-XojoCode-NextLines]
		  SqlQuery =" UPDATE   "+SrcDataObject.TableName+"   SET NAME = ? WHERE "+SrcDataObject.KeyTableName+"  = ? AND NCLLEVEL= ? "
		  ps = SrcDataObject.DBaseID.Prepare(SqlQuery)
		  ps.BindType(0, SQLitePreparedStatement.SQLITE_TEXT)
		  ps.BindType(1, SQLitePreparedStatement.SQLITE_INT64)
		  ps.BindType(2, SQLitePreparedStatement.SQLITE_INT64)
		  ps.Bind(0,NewNodeName)
		  ps.Bind(1,NodeKey)
		  ps.Bind(2,NodeLevel)
		  
		  Try
		    ps.ExecuteSQL
		    me.CellTextAt(me.SelectedRowIndex,0) = me.CellTextAt(me.SelectedRowIndex,4)+NewNodeName
		    me.CellTextAt(me.SelectedRowIndex,5) =NewNodeName
		  Catch error As DatabaseException
		    MessageBox("DB Error: " + error.Message)
		    exit sub
		  End Try
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Tree_Node_SelEmoji(NewEmoji as String)
		  Var SqlQuery as string
		  
		  Var NodeKey as Int64
		  Var NodeLevel as Int64
		  Var Parent as Int64
		  
		  // [Update-XojoCode-NextLines]
		  Var ps As SQLitePreparedStatement 
		  
		  NodeLevel = me.CellTextAt(me.SelectedRowIndex,1).ToInt64
		  NodeKey = me.CellTextAt(me.SelectedRowIndex,2).ToInt64
		  Parent = me.CellTextAt(me.SelectedRowIndex,3).ToInt64
		  
		  
		  // [Update-XojoCode-NextLines]
		  SqlQuery =" UPDATE   "+SrcDataObject.TableName+"   SET EMOJI = ? WHERE "+SrcDataObject.KeyTableName+"  = ? AND NCLLEVEL= ? "
		  ps = SrcDataObject.DBaseID.Prepare(SqlQuery)
		  ps.BindType(0, SQLitePreparedStatement.SQLITE_TEXT)
		  ps.BindType(1, SQLitePreparedStatement.SQLITE_INT64)
		  ps.BindType(2, SQLitePreparedStatement.SQLITE_INT64)
		  ps.Bind(0,NewEmoji)
		  ps.Bind(1,NodeKey)
		  ps.Bind(2,NodeLevel)
		  
		  Try
		    ps.ExecuteSQL
		    me.CellTextAt(me.SelectedRowIndex,4) =NewEmoji
		    me.CellTextAt(me.SelectedRowIndex,0) = me.CellTextAt(me.SelectedRowIndex,4)+me.CellTextAt(me.SelectedRowIndex,5) 
		    
		  Catch error As DatabaseException
		    MessageBox("DB Error: " + error.Message)
		    exit sub
		  End Try
		  
		  
		  
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		Ite_AddBrother As DesktopMenuItem
	#tag EndProperty

	#tag Property, Flags = &h0
		Ite_AddChild As DesktopMenuItem
	#tag EndProperty

	#tag Property, Flags = &h0
		Ite_DelNode As DesktopMenuItem
	#tag EndProperty

	#tag Property, Flags = &h0
		Ite_RenameNode As DesktopMenuItem
	#tag EndProperty

	#tag Property, Flags = &h0
		Ite_SelectEmoji As DesktopMenuItem
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// [Update-XojoCode-NextLines]
			// If True the checkboxes are managed by the treeview
			// If False the checkboxes are NOT managed by the treeview (they are hidden)
		#tag EndNote
		Manage_CheckBox As Boolean = True
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// [Update-XojoCode-NextLines]
			// If True the emoji are managed by the treeview
			// If False the emoji are NOT managed by the treeview (they are hidden)
		#tag EndNote
		Manage_Emoji As Boolean = True
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// Represents the data source class which will allow you to read, create and modify your data in the database
			// [Update-XojoCode-NextLines]
			
		#tag EndNote
		SrcDataObject As Class_SrcDataObject
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Width"
			Visible=true
			Group="Position"
			InitialValue="100"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Height"
			Visible=true
			Group="Position"
			InitialValue="100"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="LockLeft"
			Visible=true
			Group="Position"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="LockTop"
			Visible=true
			Group="Position"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="LockRight"
			Visible=true
			Group="Position"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="LockBottom"
			Visible=true
			Group="Position"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="TabIndex"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="TabPanelIndex"
			Visible=false
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="TabStop"
			Visible=true
			Group="Position"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="_ScrollOffset"
			Visible=false
			Group="Appearance"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="_ScrollWidth"
			Visible=false
			Group="Appearance"
			InitialValue="-1"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="AllowAutoDeactivate"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="HasBorder"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="ColumnCount"
			Visible=true
			Group="Appearance"
			InitialValue="1"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="ColumnWidths"
			Visible=true
			Group="Appearance"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="DefaultRowHeight"
			Visible=true
			Group="Appearance"
			InitialValue="-1"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Enabled"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="GridLineStyle"
			Visible=true
			Group="Appearance"
			InitialValue="0"
			Type="GridLineStyles"
			EditorType="Enum"
			#tag EnumValues
				"0 - None"
				"1 - Horizontal"
				"2 - Vertical"
				"3 - Both"
			#tag EndEnumValues
		#tag EndViewProperty
		#tag ViewProperty
			Name="HasHeader"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="HeadingIndex"
			Visible=true
			Group="Appearance"
			InitialValue="-1"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Tooltip"
			Visible=true
			Group="Appearance"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="InitialValue"
			Visible=true
			Group="Appearance"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="HasHorizontalScrollbar"
			Visible=true
			Group="Appearance"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="HasVerticalScrollbar"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="DropIndicatorVisible"
			Visible=true
			Group="Appearance"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Transparent"
			Visible=true
			Group="Appearance"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="AllowFocusRing"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Visible"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Bold"
			Visible=true
			Group="Font"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Italic"
			Visible=true
			Group="Font"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="FontName"
			Visible=true
			Group="Font"
			InitialValue="System"
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="FontSize"
			Visible=true
			Group="Font"
			InitialValue="0"
			Type="Single"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="FontUnit"
			Visible=true
			Group="Font"
			InitialValue="0"
			Type="FontUnits"
			EditorType="Enum"
			#tag EnumValues
				"0 - Default"
				"1 - Pixel"
				"2 - Point"
				"3 - Inch"
				"4 - Millimeter"
			#tag EndEnumValues
		#tag EndViewProperty
		#tag ViewProperty
			Name="Underline"
			Visible=true
			Group="Font"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="AllowAutoHideScrollbars"
			Visible=true
			Group="Behavior"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="AllowResizableColumns"
			Visible=true
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="AllowRowDragging"
			Visible=true
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="AllowRowReordering"
			Visible=true
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="AllowExpandableRows"
			Visible=true
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="RequiresSelection"
			Visible=true
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="RowSelectionType"
			Visible=true
			Group="Behavior"
			InitialValue="0"
			Type="RowSelectionTypes"
			EditorType="Enum"
			#tag EnumValues
				"0 - Single"
				"1 - Multiple"
			#tag EndEnumValues
		#tag EndViewProperty
		#tag ViewProperty
			Name="Manage_CheckBox"
			Visible=false
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Manage_Emoji"
			Visible=false
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
