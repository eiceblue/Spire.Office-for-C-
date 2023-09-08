#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ToEpub.doc";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddCoverImage.epub";

	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<DocPicture> picture = new DocPicture(doc);
	picture->LoadImageSpire(DATAPATH"/Cover.png");
	doc->SaveToEpub(outputFile.c_str(), picture);
	doc->Close();
}
