#include "pch.h"

using namespace std;
using namespace Spire::Pdf;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"AddTextStamp.pdf";
	wstring inputFile = DATAPATH"AddTextStamp.pdf";

	//Load the Pdf from disk
	intrusive_ptr<PdfDocument> document = new PdfDocument();
	document->LoadFromFile(inputFile.c_str());

	//Get the first page
	intrusive_ptr<PdfPageBase> page = document->GetPages()->GetItem(0);

	//Create a pdf template
	intrusive_ptr<PdfTemplate> template_Keyword = new PdfTemplate(125, 55);
	LPCWSTR_S ele = L"Elephant";
	intrusive_ptr<PdfTrueTypeFont> font1 = new PdfTrueTypeFont(ele, 10.f, PdfFontStyle::Italic, true);
	intrusive_ptr<PdfSolidBrush> brush = new PdfSolidBrush(new PdfRGBColor(Color::GetDarkRed()));
	intrusive_ptr<PdfPen> pen = new PdfPen(brush);
	intrusive_ptr<RectangleF> rectangle = new RectangleF(new PointF(5, 5), template_Keyword->GetSize());
	int CornerRadius = 20;
	intrusive_ptr<PdfPath> path = new PdfPath();
	path->AddArc(template_Keyword->GetBounds()->GetX(), template_Keyword->GetBounds()->GetY(), CornerRadius, CornerRadius, 180, 90);
	path->AddArc(template_Keyword->GetBounds()->GetX() + template_Keyword->GetWidth() - CornerRadius, template_Keyword->GetBounds()->GetY(), CornerRadius, CornerRadius, 270, 90);
	path->AddArc(template_Keyword->GetBounds()->GetX() + template_Keyword->GetWidth() - CornerRadius, template_Keyword->GetBounds()->GetY() + template_Keyword->GetHeight() - CornerRadius, CornerRadius, CornerRadius, 0, 90);
	path->AddArc(template_Keyword->GetBounds()->GetX(), template_Keyword->GetBounds()->GetY() + template_Keyword->GetHeight() - CornerRadius, CornerRadius, CornerRadius, 90, 90);
	path->AddLine(template_Keyword->GetBounds()->GetX(), template_Keyword->GetBounds()->GetY() + template_Keyword->GetHeight() - CornerRadius, template_Keyword->GetBounds()->GetX(), template_Keyword->GetBounds()->GetY() + CornerRadius / 2);
	template_Keyword->GetGraphics()->DrawPath(pen, path);

	//Draw stamp text
	std::wstring s1 = L"REVISED\n";
	LPCWSTR_S t = L"MM dd, yyyy";
	wstring timeString = DateTime::GetNow()->ToString(t);
	wstring s2 = L"by E-iceblue at " + timeString;
	template_Keyword->GetGraphics()->DrawString(s1.c_str(), font1, brush, new PointF(5, 10));
	LPCWSTR_S luc = L"Lucida Sans Unicod";
	intrusive_ptr<PdfTrueTypeFont> font2 = new PdfTrueTypeFont(luc, 9.f, PdfFontStyle::Bold, true);

	template_Keyword->GetGraphics()->DrawString(s2.c_str(), font2, brush, new PointF(2, 30));

	//Create a rubber stamp
	intrusive_ptr<PdfRubberStampAnnotation> stamp = new PdfRubberStampAnnotation(rectangle);
	intrusive_ptr<PdfAppearance> apprearance = new PdfAppearance(stamp);
	apprearance->SetNormal(template_Keyword);
	stamp->SetAppearance(apprearance);

	//Draw stamp into page
	page->GetAnnotationsWidget()->Add(stamp);

	//Save the document
	document->SaveToFile(outputFile.c_str());
	document->Close();
}
