#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;


void CompressContent(intrusive_ptr<PdfDocument> doc)
{
	//Disable the incremental update
	doc->GetFileInfo()->SetIncrementalUpdate(false);

	//Set the compression level to best
	doc->SetCompressionLevel(PdfCompressionLevel::Best);
}

void CompressImage(intrusive_ptr<PdfDocument> doc)
{
	//Disable the incremental update
	doc->GetFileInfo()->SetIncrementalUpdate(false);

	//Traverse all pages
	for (int i = 0; i < doc->GetPages()->GetCount(); i++)
	{
		intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(i);
		if (page != nullptr)
		{
			if (page->GetImagesInfo().size() != 0)
			{
				vector<intrusive_ptr<PdfImageInfo>> imageInfo = page->GetImagesInfo();
				for (int j = 0; j < imageInfo.size(); j++)
				{
					page->TryCompressImage(imageInfo[i]->GetIndex());
				}
			}
		}
	}
}

int main()
{
	wstring inputFile = DATAPATH"CompressDocument.pdf";
	wstring outputFile = OUTPUTPATH"CompressDocument.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();

	doc->LoadFromFile(inputFile.c_str());
	//Compress the content in document
	CompressContent(doc);

	//Compress the images in document
	CompressImage(doc);
	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
