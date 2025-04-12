#tag Class
Protected Class App
Inherits DesktopApplication
	#tag Method, Flags = &h0
		Sub CloseDatabaseFile()
		  Dim message as string
		  
		  if not (MainDB_State="CLOSE") then
		    message =  "Closing the database  : "+EndOfLine+MainDB.DatabaseFile.Name
		    MainDB.Close
		    MessageBox message
		    MainDB=new SQLiteDatabase
		    MainDB_State="CLOSE"
		  else
		    MessageBox  "No open database !"
		  end if
		  
		  Exception err as NilObjectException
		    MainDB_State="CLOSE"
		    MessageBox "Error : No open database !"
		    
		    
		    
		    
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CreateDatabaseFile() As Boolean
		  Dim MyXojoDBDict as new Class_XojoDBDict
		  
		  
		  Try
		    // Create the database file
		    MainDB.CreateDatabase 
		    MyXojoDBDict.DBaseID=App.MainDB
		    MyXojoDBDict.Initialise_Base
		    return App.OpenDatabaseFile
		  Catch error As IOException
		    MessageBox("The database file could not be created: " + error.Message)
		    return false
		  End Try
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function OpenDatabaseFile() As Boolean
		  
		  Try
		    MainDB.Connect 
		    MainDB_State="OPEN"
		    return true
		  Catch error As DatabaseException
		    MessageBox("Error connecting to the database: " +EndOfLine+ error.Message)
		    return false
		  End Try
		End Function
	#tag EndMethod


	#tag Note, Name = ReadMe: Adapting to your needs
		I wrote in the comments a tag : // [Update-XojoCode-NextLines]
		
		When you see this tag in comments, it means that the program probably needs to be modified in the following lines to meet your needs.
		If you see this tag in the declaration of a variable, it means that you will probably need to modify the name of the variable or its default value to adapt the program to your needs.
		Please note: This does not mean that depending on what you want to do, you will only need to modify these lines.
		This is simply an aid to help you locate important points that most likely need to be modified to suit your needs. 
		I do not claim to be exhaustive on this subject.
		
		
		This module is based on the management of an SQLite Database Table :
		
		CREATE TABLE NOMENCLATURE (          
		    KEY_ID            BIGINT PRIMARY KEY,      
		    NCLLEVEL      BIGINT NOT NULL,      
		    NCLPARENT   BIGINT,      
		    NAME              VARCHAR(50), 
		    IS_SELECTED INTEGER DEFAULT 0, 
		    EMOJI              VARCHAR(1) DEFAULT '📁'
		);
		
	#tag EndNote


	#tag Property, Flags = &h0
		MainDB As SQLiteDatabase
	#tag EndProperty

	#tag Property, Flags = &h0
		MainDB_State As String = "CLOSE"
	#tag EndProperty


	#tag Constant, Name = kEditClear, Type = String, Dynamic = False, Default = \"&Delete", Scope = Public
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"&Delete"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"&Delete"
	#tag EndConstant

	#tag Constant, Name = kFileQuit, Type = String, Dynamic = False, Default = \"&Quit", Scope = Public
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"E&xit"
	#tag EndConstant

	#tag Constant, Name = kFileQuitShortcut, Type = String, Dynamic = False, Default = \"", Scope = Public
		#Tag Instance, Platform = Mac OS, Language = Default, Definition  = \"Cmd+Q"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"Ctrl+Q"
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=false
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=false
			Group="ID"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=false
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=false
			Group="Position"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=false
			Group="Position"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="AllowAutoQuit"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="AllowHiDPI"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="BugVersion"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Copyright"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Description"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LastWindowIndex"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="MajorVersion"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="MinorVersion"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="NonReleaseVersion"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="RegionCode"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="StageCode"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Version"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="_CurrentEventTime"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="MainDB_State"
			Visible=false
			Group="Behavior"
			InitialValue="CLOSE"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
