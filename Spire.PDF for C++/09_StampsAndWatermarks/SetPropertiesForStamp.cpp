#include "pch.h"

using namespace std;
using namespace Spire::Pdf;
using namespace Spire::Common;

int main()
{
	wstring inputFile = DATAPATH"TextStamp.pdf";
	wstring outputFile = OUTPUTPATH"SetPropertiesForStamp.pdf";

	//Load old PDF from disk.
	intrusive_ptr<PdfDocument> pdf = new PdfDocument();
	pdf->LoadFromFile(inputFile.c_str());
	//Get the first page
	intrusive_ptr<PdfPageBase> page = pdf->GetPages()->GetItem(0);
	//Traverse annotations widget
	for (int i = 0; i < page->GetAnnotationsWidget()->GetCount(); i++)
	{
		intrusive_ptr<PdfAnnotation> annotation = page->GetAnnotationsWidget()->GetItem(i);
		wchar_t nm_w[100];
		swprintf(nm_w, 100, L"%hs", typeid(PdfRubberStampAnnotationWidget).name());
		LPCWSTR_S newName = nm_w;
		if (wcscmp(newName, annotation->GetInstanceTypeName()) == 0)
		{
			intrusive_ptr<PdfRubberStampAnnotationWidget> stamp = Object::Convert<PdfRubberStampAnnotationWidget>(annotation);
			stamp->SetAuthor(L"TestUser");
			stamp->SetSubject(L"E-iceblue");
			stamp->SetCreationDate(DateTime::GetNow());
			stamp->SetModifiedDate(DateTime::GetNow());
		}
	}

	//Save the document
	pdf->SaveToFile(outputFile.c_str());
	pdf->Close();
}
