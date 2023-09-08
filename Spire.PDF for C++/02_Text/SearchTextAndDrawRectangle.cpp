#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
int main() {

	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SearchReplaceTemplate.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"earchTextAndDrawRectangle.pdf";
	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	intrusive_ptr<PdfTextFindCollection> collection = page->FindText(L"Spire.PDF for .NET", Find_TextFindParameter::IgnoreCase);
	for(intrusive_ptr<PdfTextFind> find : collection->GetFinds())
	{
		// Draw a rectangle with red pen
		page->GetCanvas()->DrawRectangle(new PdfPen(PdfBrushes::GetRed(), 0.9f), find->GetBounds());
	}
	doc->SaveToFile(outputFile.c_str());
	doc->Close();

}

