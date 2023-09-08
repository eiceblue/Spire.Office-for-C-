#include "../pch.h"


using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main()
{
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"RadioButton.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddRadioButtonCaption.pdf";

	intrusive_ptr<PdfDocument> pdf = new PdfDocument();
	pdf->LoadFromFile(inputFile.c_str());

	//Get pdf forms
	intrusive_ptr<PdfFormWidget> formWidget = Object::Convert<PdfFormWidget>(pdf->GetForm());

	//Find the radio button field and add capture
	for (int i = 0; i < formWidget->GetFieldsWidget()->GetCount(); i++)
	{
		intrusive_ptr<PdfField> field = formWidget->GetFieldsWidget()->GetItem(i);
		wchar_t nm_w[100];
		swprintf(nm_w, 100, L"%hs", typeid(PdfRadioButtonListFieldWidget).name());
		LPCWSTR_S newName = nm_w;
		if (wcscmp(newName, field->GetInstanceTypeName()) == 0)
		{
			intrusive_ptr<PdfRadioButtonListFieldWidget> radioButton = Object::Convert<PdfRadioButtonListFieldWidget>(field);
			wstring str = radioButton->GetName();
			if (str == L"RadioButton")
			{
				//Get the page
				intrusive_ptr<PdfPageBase> page = radioButton->GetPage();

				//Set capture name
				std::wstring text = L"Radio button caption";
				//Set font, pen and brush
				intrusive_ptr<PdfFont> font = new PdfFont(PdfFontFamily::Helvetica, 12.0f);
				intrusive_ptr<PdfPen> pen = new PdfPen(new PdfRGBColor(Color::GetRed()), 0.02f);
				intrusive_ptr<PdfSolidBrush> brush = new PdfSolidBrush(new PdfRGBColor(Color::GetRed()));
				//Set the capture location
				float x = radioButton->GetLocation()->GetX();
				float y = radioButton->GetLocation()->GetY() - font->MeasureString(text.c_str())->GetHeight() - 10;
				;
				//Draw capture
				page->GetCanvas()->DrawString(text.c_str(), font, pen, brush, x, y);

			}
		}
	}
	//Save the document
	pdf->SaveToFile(outputFile.c_str());
	pdf->Close();
}