#include "pch.h"
using namespace Spire::Doc;
#define stringify(name) # name
wstring getFormFieldType(FormFieldType type)
{
	switch (type)
	{
	case FormFieldType::CheckBox:
		return L"CheckBox";
		break;
	case FormFieldType::DropDown:
		return L"DropDown";
		break;
	case FormFieldType::TextInput:
		return L"TextInput";
		break;
	case FormFieldType::Unknown:
		return L"Unknown";
		break;
	}
}
int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"FillFormField.doc";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetFormFieldByName.txt";

	wstring* sb = new wstring();

	//Open a Word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

	//Get form field by name
	intrusive_ptr<FormField> formField = section->GetBody()->GetFormFields()->GetItem(L"email");
	wstring formFieldName = formField->GetName();
	wstring formFieldNameType = getFormFieldType(formField->GetFormFieldType());
	sb->append(L"The name of the form field is " + formFieldName + L"\n");
	sb->append(L"The type of the form field is " + formFieldNameType);

	wofstream out;
	out.open(outputFile);
	out.flush();
	out << sb->c_str();
	out.close();

	document->Close();
	delete sb;
}


