#include "../pch.h"

using namespace std;
using namespace Spire::Pdf;

int main() {
	wstring fn1 = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DeleteImageByBounds.pdf";

	
	//Open pdf document
	intrusive_ptr<PdfDocument> pdf = new PdfDocument();
	pdf->LoadFromFile((fn1 + L"DeleteImageByBounds.pdf").c_str());

	//Get the first page
	intrusive_ptr<PdfPageBase> page = pdf->GetPages()->GetItem(0);
	//Get the information of all images in this page 
	vector<intrusive_ptr<PdfImageInfo>> imageInfo = page->GetImagesInfo();
	for (int i = 0; i < imageInfo.size(); i++)
	{

		if (imageInfo[i]->GetBounds()->Contains(49.68f, 73.1f))
		{
			page->DeleteImage(imageInfo[i]);
		}

		//Case 2:  
		if (imageInfo[i]->GetBounds()->IntersectsWith(new RectangleF(100.0f, 500.0f, 30.0f, 40.0f))) {
			page->DeleteImage(imageInfo[i]);
		}

	}
	//save to file
	pdf->SaveToFile(outputFile.c_str());
	pdf->Close();
}









