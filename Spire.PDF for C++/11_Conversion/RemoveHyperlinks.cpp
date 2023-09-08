#include "pch.h"

using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring outputFile = OUTPUTPATH"RemoveHyperlinks.pdf";
	wstring inputFile = DATAPATH"RemoveHyperlinks.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	intrusive_ptr<PdfAnnotationCollection> widgetCollection = page->GetAnnotationsWidget();
	//Verify whether widgetCollection is null or not
	if (widgetCollection->GetCount() > 0) {
		for (int i = widgetCollection->GetCount() - 1; i >= 0; i--) {
			intrusive_ptr<PdfAnnotation> annotation = widgetCollection->GetItem(i);
			wchar_t nm_w[100];
			swprintf(nm_w, 100, L"%hs", typeid(PdfTextWebLinkAnnotationWidget).name());
			LPCWSTR_S newName = nm_w;
			if (wcscmp(newName, annotation->GetInstanceTypeName()) == 0) {
				intrusive_ptr<PdfTextWebLinkAnnotationWidget> link = Object::Convert<PdfTextWebLinkAnnotationWidget>(annotation);
				widgetCollection->Remove(link);
			}
		}
	}
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
