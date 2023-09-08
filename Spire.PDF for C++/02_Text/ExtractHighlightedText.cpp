#include "../pch.h"

using namespace Spire::Pdf;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path +L"ExtractHighlightedText.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path +L"ExtractHighlightedText.txt";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	intrusive_ptr<PdfTextMarkupAnnotationWidget> annotation;
	wstring str = L"Extracted hightlighted text:\n";
	for (int i = 0; i < page->GetAnnotationsWidget()->GetCount(); i++) {
		wchar_t nm_w[100];
		swprintf(nm_w, 100, L"%hs", typeid(PdfTextMarkupAnnotationWidget).name());
		LPCWSTR_S newName = nm_w;
		if (wcscmp(newName, page->GetAnnotationsWidget()->GetItem(i)->GetInstanceTypeName()) == 0)
		{
			annotation = Object::Convert<PdfTextMarkupAnnotationWidget>(page->GetAnnotationsWidget()->GetItem(i));
			str += page->ExtractText(annotation->GetBounds());
			intrusive_ptr<PdfRGBColor> color = annotation->GetTextMarkupColor();
		}
	}
	wofstream os;
	os.open(outputFile, ios::trunc);
	os << str;
	os.close();
	doc->Close();
}

