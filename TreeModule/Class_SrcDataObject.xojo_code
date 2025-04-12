#tag Class
Protected Class Class_SrcDataObject
	#tag Method, Flags = &h0
		Sub Constructor()
		  Fields_Init
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Fields_Init()
		  // Initializes class properties with default values
		  
		  // [Update-XojoCode-NextLines]
		  KeyTableValue         = 0
		  Field_Name             = ""
		  Field_NclLevel          = 0
		  Field_NclParent        = 0
		  Field_IsSelected       = 0
		  
		  //Warning : You must leave ONE and only ONE character, do not put empty characters!
		  Field_Emoji              ="📁" 
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub Fields_Load()
		  // Loads the values ​​of a table record into the corresponding properties of the class.
		  
		  // [Update-XojoCode-NextLines]
		  Try
		    KeyTableValue         = DBaseRS.Column(KeyTableName).Int64Value
		  Catch error As NilObjectException
		    KeyTableValue    = 0
		  End Try
		  
		  Try
		    Field_Name      = DBaseRS.Column("NAME").StringValue
		  Catch error As NilObjectException
		    Field_Name    = ""
		  End Try
		  
		  
		  Try
		    Field_NclLevel = DBaseRS.Column("NCLLEVEL").Int64Value
		  Catch error As NilObjectException
		    Field_NclLevel    = 0
		  End Try
		  
		  Try
		    Field_NclParent  = DBaseRS.Column("NCLPARENT").Int64Value
		  Catch error As NilObjectException
		    Field_NclParent    = 0
		  End Try
		  
		  Try
		    Field_IsSelected = DBaseRS.Column("IS_SELECTED").Int64Value
		  Catch error As NilObjectException
		    Field_IsSelected    = 0
		  End Try
		  
		  Try
		    Field_Emoji = DBaseRS.Column("EMOJI").StringValue
		  Catch error As NilObjectException
		    Field_Emoji = "📁"
		  End Try
		  
		  
		  
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Find_FreeKeyTableValue() As int64
		  // The purpose of this function is to generate a primary key value for the table record
		  
		  Dim rs         as RowSet
		  
		  rs = DBaseID.SelectSQL("SELECT IFNULL(MAX("+KeyTableName+")+1,1) AS MAX_KEY_VALUE FROM "+TableName)
		  
		  if  not (rs=NIL) then
		    rs.MoveToFirstRow
		    KeyTableValue        = rs.Column("MAX_KEY_VALUE").Value
		  else
		    KeyTableValue        = 1
		  end if
		  
		  return KeyTableValue
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Record_Create(Create_KeyTableValue as Boolean = True) As Boolean
		  // Create a new record with class properties values.
		  
		  // By default the key is incremented, but you can force the key to a given value by update KeyTableValue propertie.
		  if Create_KeyTableValue= True then KeyTableValue=Find_FreeKeyTableValue()
		  
		  Var row As New DatabaseRow
		  
		  // [Update-XojoCode-NextLines]
		  row.Column(KeyTableName).Int64Value     = KeyTableValue
		  row.Column("NAME").StringValue            = Field_Name
		  row.Column("NCLLEVEL").Int64Value        = Field_NclLevel
		  row.Column("NCLPARENT").Int64Value    = Field_NclParent
		  row.Column("IS_SELECTED").Int64Value    = Field_IsSelected
		  row.Column("EMOJI").StringValue           = Field_Emoji
		  
		  Try
		    DBaseID.AddRow(TableName, row)
		  Catch error As DatabaseException
		    MessageBox("DB Error: " + error.Message)
		    return false
		  End Try
		  
		  return true
		  
		  
		  
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Record_Delete(Del_KeySelection as int64 = -1) As Boolean
		  
		  Select Case Del_KeySelection
		    
		  Case Is >=0
		    
		    // if Del_KeySelection>=0 then you want to delete a specific record and not the current record
		    Try
		      DBaseID.ExecuteSQL("DELETE FROM "+TableName+" WHERE "+KeyTableName+" = "+Del_KeySelection.ToString)
		      return true
		    Catch error As DatabaseException
		      MessageBox("DB Error: " + error.Message)
		      return false
		    End Try
		    
		  Case else
		    
		    // if Del_KeySelection=-1 then Delete current record 
		    Try
		      DBaseRS.RemoveRow
		      return true
		    Catch error As DatabaseException
		      MessageBox("DB Error: " + error.Message)
		      return false
		    End Try
		    
		  end Select
		  
		  
		  return false
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Record_ReadFirst() As Boolean
		  // Read the first record of the sql query
		  
		  Fields_Init
		  
		  
		  Try
		    DBaseRS.MoveToFirstRow
		  Catch error As DatabaseException
		    MessageBox("DB Error: " + error.Message)
		    return false
		  End Try
		  
		  if  DBaseRS.BeforeFirstRow or DBaseRS.AfterLastRow  then return false
		  
		  
		  Fields_Load()
		  
		  return true
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Record_ReadNext() As Boolean
		  // Read the next record of the sql query
		  
		  Fields_Init
		  
		  Try
		    DBaseRS.MoveToNextRow
		  Catch error As DatabaseException
		    MessageBox("DB Error: " + error.Message)
		    return false
		  End Try
		  
		  if  DBaseRS.AfterLastRow  then return false
		  
		  
		  Fields_Load()
		  
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Record_ReadPrevious() As Boolean
		  // Read the previous record of the sql query
		  
		  Fields_Init
		  
		  Try
		    DBaseRS.MoveToPreviousRow
		  Catch error As DatabaseException
		    MessageBox("DB Error: " + error.Message)
		    return false
		  End Try
		  
		  if  DBaseRS.BeforeFirstRow  then return false
		  
		  
		  Fields_Load()
		  
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Record_Update() As Boolean
		  // Update a record with class properties
		  
		  Try
		    DBaseRS.EditRow
		    
		    // [Update-XojoCode-NextLines]
		    DBaseRS.Column(KeyTableName).Int64Value  = KeyTableValue
		    DBaseRS.Column("NAME").StringValue            = Field_Name
		    DBaseRS.Column("NCLLEVEL").Int64Value        = Field_NclLevel
		    DBaseRS.Column("NCLPARENT").Int64Value   = Field_NclParent
		    DBaseRS.Column("IS_SELECTED").Int64Value   = Field_IsSelected
		    DBaseRS.Column("EMOJI").StringValue           = Field_Emoji
		    
		    DBaseRS.SaveRow
		    
		  Catch error As DatabaseException
		    MessageBox("DB Error: " + error.Message)
		    return false
		  End Try
		  
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Run_SqlQuerySource() As Boolean
		  // Execute the query which will be the data source of the class
		  
		  
		  
		  Try
		    DBaseRS=DBaseID.SelectSQL(SqlQuerySource,SqlParams())
		  Catch error As DatabaseException
		    MessageBox("Bad Sql resquest: " + error.Message)
		    return false
		  End Try
		  
		  
		  Try
		    DBaseRS.MoveToFirstRow
		  Catch error As DatabaseException
		    MessageBox("DB Error: " + error.Message)
		    return false
		  End Try
		  
		  Fields_Load()
		  
		  return true
		  
		  
		  
		  
		  
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		BLOCAGE As String = "N"
	#tag EndProperty

	#tag Property, Flags = &h0
		DBaseID As SQLiteDatabase
	#tag EndProperty

	#tag Property, Flags = &h0
		DBaseRS As RowSet
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// [Update-XojoCode-NextLines]
		#tag EndNote
		Field_Emoji As String = "📁"
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// [Update-XojoCode-NextLines]
		#tag EndNote
		Field_IsSelected As Int64 = 0
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// [Update-XojoCode-NextLines]
		#tag EndNote
		Field_Name As String
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// [Update-XojoCode-NextLines]
		#tag EndNote
		Field_NclLevel As Int64 = 0
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// [Update-XojoCode-NextLines]
		#tag EndNote
		Field_NclParent As Int64 = 0
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// Name of the key identifying the record in the table
			// [Update-XojoCode-NextLines]
		#tag EndNote
		KeyTableName As string = "KEY_ID"
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// [Update-XojoCode-NextLines]
		#tag EndNote
		KeyTableValue As Int64 = -1
	#tag EndProperty

	#tag Property, Flags = &h0
		SqlParams() As Variant
	#tag EndProperty

	#tag Property, Flags = &h0
		SqlQuerySource As String
	#tag EndProperty

	#tag Property, Flags = &h0
		#tag Note
			// Name of the table
			// [Update-XojoCode-NextLines]
		#tag EndNote
		TableName As string = "NOMENCLATURE"
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="BLOCAGE"
			Visible=false
			Group="Behavior"
			InitialValue="N"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
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
			Name="Name"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Field_NclLevel"
			Visible=false
			Group="Behavior"
			InitialValue="1"
			Type="Int64"
			EditorType="MultiLineEditor"
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
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="KeyTableValue"
			Visible=false
			Group="Behavior"
			InitialValue="-1"
			Type="Int64"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Field_Name"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="KeyTableName"
			Visible=false
			Group="Behavior"
			InitialValue="REFNUM"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="TableName"
			Visible=false
			Group="Behavior"
			InitialValue="CUSTOMERS"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="SqlQuerySource"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Field_NclParent"
			Visible=false
			Group="Behavior"
			InitialValue="1"
			Type="Int64"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Field_IsSelected"
			Visible=false
			Group="Behavior"
			InitialValue="0"
			Type="Int64"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Field_Emoji"
			Visible=false
			Group="Behavior"
			InitialValue="📁"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
