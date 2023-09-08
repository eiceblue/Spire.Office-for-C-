#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path+L"SearchReplaceTemplate.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile= output_path+L"SearchTextAndAddHyperlink.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	intrusive_ptr<PdfTextFindCollection> collection = page->FindText(L"e-iceblue", Find_TextFindParameter::IgnoreCase);
	wstring url = L"http://www.e-iceblue.com";
	for(intrusive_ptr<PdfTextFind> find : collection->GetFinds())
	{
		// Create a PdfUriAnnotation object to add hyperlink for the searched text 
		intrusive_ptr<PdfUriAnnotation> uri = new PdfUriAnnotation(find->GetBounds());
		uri->SetUri(url.c_str());
		uri->SetBorder(new PdfAnnotationBorder(1.f));
		uri->SetColor(new PdfRGBColor(Color::GetBlue()));
		page->GetAnnotationsWidget()->Add(uri);
	}
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}

