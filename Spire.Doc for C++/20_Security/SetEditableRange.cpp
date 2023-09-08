#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SetEditableRange.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetEditableRange.docx";

	//Create a new document
	intrusive_ptr<Document> document = new Document();
	//Load file from disk
	document->LoadFromFile(inputFile.c_str());
	//Protect whole document
	document->Protect(ProtectionType::AllowOnlyReading, L"password");
	//Create tags for permission start and end
	intrusive_ptr<PermissionStart> start = new PermissionStart(document, L"testID");
	intrusive_ptr<PermissionEnd> end = new PermissionEnd(document, L"testID");
	//Add the start and end tags to allow the first paragraph to be edited.
	document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0)->GetChildObjects()->Insert(0, start);
	document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0)->GetChildObjects()->Add(end);
	//Save the document
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}