#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"HeaderAndFooter.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"LockHeader.docx";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

	//Protect the document and set the ProtectionType as AllowOnlyFormFields
	doc->Protect(ProtectionType::AllowOnlyFormFields, L"123");

	//Set the ProtectForm as false to unprotect the section
	section->SetProtectForm(false);


	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
