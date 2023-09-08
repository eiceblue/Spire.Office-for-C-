#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Pdf_3.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetAllAnnotationsFromPage.txt";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfAnnotationCollection> annotations = doc->GetPages()->GetItem(0)->GetAnnotationsWidget();
	wstring content;
	for (int i = 0; i < annotations->GetCount(); i++) {
		wchar_t nm_w[100];
		swprintf(nm_w, 100, L"%hs", typeid(PdfPopupAnnotationWidget).name());
		LPCWSTR_S newName = nm_w;
		if (wcscmp(newName, annotations->GetItem(i)->GetInstanceTypeName()) == 0)
			continue;
		content.append(L"Annotation information: \n");
		content.append(L"Text: ");
		content.append(annotations->GetItem(i)->GetText());
		wstring modifiedDate = annotations->GetItem(i)->GetModifiedDate()->ToString();
		content.append(L"\nModifiedDate: " + modifiedDate);
		content.append(L"\n");
	}
	wofstream os;
	os.open(outputFile, ios::trunc);
	os << content;
	os.close();
	doc->Close();
}


