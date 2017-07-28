Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Text
Imports System.Text.RegularExpressions

Namespace com.bodurov
	Public NotInheritable Class DataSourceCreator
		Private Sub New()
		End Sub
		Private Shared ReadOnly PropertNameRegex As New Regex("^[A-Za-z]+[A-Za-z0-9_]*$", RegexOptions.Singleline)

		Private Shared ReadOnly _typeBySigniture As New Dictionary(Of String, Type)()

		Public Shared Function ToDataSource(list As IEnumerable(Of IDictionary)) As IEnumerable
			Dim firstDict As IDictionary = Nothing
			Dim hasData As Boolean = False
			For Each currentDict As IDictionary In list
				hasData = True
				firstDict = currentDict
				Exit For
			Next
			If Not hasData Then
				Return New Object() {}
			End If
			If firstDict Is Nothing Then
				Throw New ArgumentException("IDictionary entry cannot be null")
			End If

			Dim typeSignature As String = GetTypeSignature(firstDict, list)

			Dim objectType As Type = GetTypeByTypeSignature(typeSignature)

			If objectType Is Nothing Then
				Dim tb As TypeBuilder = GetTypeBuilder(typeSignature)

				Dim constructor As ConstructorBuilder = tb.DefineDefaultConstructor(MethodAttributes.[Public] Or MethodAttributes.SpecialName Or MethodAttributes.RTSpecialName)


				For Each pair As DictionaryEntry In firstDict
					If PropertNameRegex.IsMatch(Convert.ToString(pair.Key), 0) Then

						If pair.Value Is Nothing Then

							CreateProperty(tb, Convert.ToString(pair.Key), GetValueType(pair.Value, list, pair.Key))
						End If
					Else
						Throw New ArgumentException("Each key of IDictionary must be " & vbCr & vbLf & "                                alphanumeric and start with character.")
					End If
				Next
				objectType = tb.CreateType()

				_typeBySigniture.Add(typeSignature, objectType)
			End If

			Return GenerateEnumerable(objectType, list, firstDict)
		End Function



		Private Shared Function GetTypeByTypeSignature(typeSigniture As String) As Type
			Dim type As Type
			Return If(_typeBySigniture.TryGetValue(typeSigniture, type), type, Nothing)
		End Function

		Private Shared Function GetValueType(value As Object, list As IEnumerable(Of IDictionary), key As Object) As Type
			If value Is Nothing Then
				For Each dictionary As IDictionary In list
					If dictionary.Contains(key) Then
						value = dictionary(key)
						If value IsNot Nothing Then
							Exit For
						End If
					End If
				Next
			End If
			Return If((value Is Nothing), GetType(Object), value.[GetType]())
		End Function

		Private Shared Function GetTypeSignature(firstDict As IDictionary, list As IEnumerable(Of IDictionary)) As String
			Dim sb = New StringBuilder()
			For Each pair As DictionaryEntry In firstDict
				sb.AppendFormat("_{0}_{1}", pair.Key, GetValueType(pair.Value, list, pair.Key))
			Next
			Return sb.ToString().GetHashCode().ToString().Replace("-", "Minus")
		End Function

        Private Shared Function GenerateEnumerable(ByVal objectType As Type, ByVal list As IEnumerable(Of IDictionary), ByVal firstDict As IDictionary) As IEnumerable
			Dim listType As Type = GetType(List(Of )).MakeGenericType(New Type() {objectType})
			Dim listOfCustom As IList = Activator.CreateInstance(listType)

			For Each currentDict As IDictionary In list
				If currentDict Is Nothing Then
					Throw New ArgumentException("IDictionary entry cannot be null")
				End If
				Dim row = Activator.CreateInstance(objectType)
				For Each pair As DictionaryEntry In firstDict
					If currentDict.Contains(pair.Key) Then
						Dim [property] As PropertyInfo = objectType.GetProperty(Convert.ToString(pair.Key))

						Dim value = currentDict(pair.Key)
						If value IsNot Nothing AndAlso value.[GetType]() <> [property].PropertyType AndAlso Not [property].PropertyType.IsGenericType Then
							Try
								value = Convert.ChangeType(currentDict(pair.Key), [property].PropertyType, Nothing)
							Catch
							End Try
						End If

						[property].SetValue(row, value, Nothing)
					End If
				Next
				listType.GetMethod("Add").Invoke(listOfCustom, New Object() {row})
			Next
			Return TryCast(listOfCustom, IEnumerable)
		End Function

		Private Shared Function GetTypeBuilder(typeSigniture As String) As TypeBuilder
			Dim an As New AssemblyName(Convert.ToString("TempAssembly") & typeSigniture)
			Dim assemblyBuilder As AssemblyBuilder = AppDomain.CurrentDomain.DefineDynamicAssembly(an, AssemblyBuilderAccess.Run)
			Dim moduleBuilder As ModuleBuilder = assemblyBuilder.DefineDynamicModule("MainModule")

			Dim tb As TypeBuilder = moduleBuilder.DefineType(Convert.ToString("TempType") & typeSigniture, TypeAttributes.[Public] Or TypeAttributes.[Class] Or TypeAttributes.AutoClass Or TypeAttributes.AnsiClass Or TypeAttributes.BeforeFieldInit Or TypeAttributes.AutoLayout, GetType(Object))
			Return tb
		End Function

		Private Shared Sub CreateProperty(tb As TypeBuilder, propertyName As String, propertyType As Type)
			If propertyType.IsValueType AndAlso Not propertyType.IsGenericType Then
				propertyType = GetType(Nullable(Of )).MakeGenericType(propertyType)
			End If

			Dim fieldBuilder As FieldBuilder = tb.DefineField(Convert.ToString("_") & propertyName, propertyType, FieldAttributes.[Private])


			Dim propertyBuilder As PropertyBuilder = tb.DefineProperty(propertyName, PropertyAttributes.HasDefault, propertyType, Nothing)
			Dim getPropMthdBldr As MethodBuilder = tb.DefineMethod(Convert.ToString("get_") & propertyName, MethodAttributes.[Public] Or MethodAttributes.SpecialName Or MethodAttributes.HideBySig, propertyType, Type.EmptyTypes)

			Dim getIL As ILGenerator = getPropMthdBldr.GetILGenerator()

			getIL.Emit(OpCodes.Ldarg_0)
			getIL.Emit(OpCodes.Ldfld, fieldBuilder)
			getIL.Emit(OpCodes.Ret)

			Dim setPropMthdBldr As MethodBuilder = tb.DefineMethod(Convert.ToString("set_") & propertyName, MethodAttributes.[Public] Or MethodAttributes.SpecialName Or MethodAttributes.HideBySig, Nothing, New Type() {propertyType})

			Dim setIL As ILGenerator = setPropMthdBldr.GetILGenerator()

			setIL.Emit(OpCodes.Ldarg_0)
			setIL.Emit(OpCodes.Ldarg_1)
			setIL.Emit(OpCodes.Stfld, fieldBuilder)
			setIL.Emit(OpCodes.Ret)

			propertyBuilder.SetGetMethod(getPropMthdBldr)
			propertyBuilder.SetSetMethod(setPropMthdBldr)
		End Sub
	End Class
End Namespace
