#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SampleB_2.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddBarcodeImage.docx";

	//Open a Word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	wstring imgPath = input_path + L"barcode.png";

	//Add barcode image
	intrusive_ptr<DocPicture> picture = document->GetSections()->GetItemInSectionCollection(0)->AddParagraph()->AppendPicture(imgPath.c_str());

	//Save to file
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
