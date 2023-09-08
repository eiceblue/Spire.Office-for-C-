#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SampleB_2.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ReplaceTextWithMergeField.docx";

	//Open a Word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Find the text that will be replaced
	intrusive_ptr<TextSelection> ts = document->FindString(L"Test", true, true);

	intrusive_ptr<TextRange> tr = ts->GetAsOneRange();

	//Get the paragraph
	intrusive_ptr<Paragraph> par = tr->GetOwnerParagraph();

	//Get the index of the text in the paragraph
	int index = par->GetChildObjects()->IndexOf(tr);

	//Create a new field
	intrusive_ptr<MergeField> field = new MergeField(document);
	field->SetFieldName(L"MergeField");

	//Insert field at specific position
	par->GetChildObjects()->Insert(index, field);

	//Remove the text
	par->GetChildObjects()->Remove(tr);

	//Save to file
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}