#include "pch.h"
using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TableTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddPictureToTableCell.docx";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first table from the first section of the document
	intrusive_ptr<Table> table1 = Object::Dynamic_cast<Table>(doc->GetSections()->GetItemInSectionCollection(0)->GetTables()->GetItemInTableCollection(0));

	intrusive_ptr<DocPicture> picture = table1->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(2)->GetParagraphs()->GetItemInParagraphCollection(0)->AppendPicture(DATAPATH"/Spire.Doc.png");
	
	picture->SetWidth(100);
	picture->SetHeight(100);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
