#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Attachment.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"Attachment.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);

	float y = 320;

	//Title
	intrusive_ptr<PdfBrush> brush1 = PdfBrushes::GetCornflowerBlue();
	intrusive_ptr<PdfTrueTypeFont> font1 = new PdfTrueTypeFont(L"Arial", 18.f, PdfFontStyle::Bold, true);
	intrusive_ptr<PdfStringFormat> format1 = new PdfStringFormat(PdfTextAlignment::Center);
	wstring text1 = L"Attachment";
	page->GetCanvas()->DrawString(text1.c_str(), font1, brush1, page->GetCanvas()->GetClientSize()->GetWidth() / 2, y, format1);
	y = y + font1->MeasureString(text1.c_str(), format1)->GetHeight();
	y = y + 10;

	//Add an attachment
	intrusive_ptr<PdfAttachment> attachment = new PdfAttachment(L"Header.png");
	wstring inputFile_h = input_path + L"Header.png";
	ifstream is(inputFile_h.c_str(), ifstream::in | ios::binary);

	is.seekg(0, is.end);
	int length = is.tellg();
	is.seekg(0, is.beg);

	char* buffer = new  char[length];

	is.read(buffer, length);

	attachment->SetData((unsigned char*)buffer, length);

	attachment->SetDescription(L"Page header picture of demo.");
	attachment->SetMimeType(L"image/png");
	doc->GetAttachments()->Add(attachment);

	//Add an attachment
	attachment = new PdfAttachment(L"Footer.png");
	wstring inputFile_f = input_path + L"Footer.png";
	ifstream is5(inputFile_f.c_str(), ifstream::in | ios::binary);

	is5.seekg(0, is5.end);
	int length5 = is5.tellg();
	is5.seekg(0, is5.beg);

	char* buffer5 = new  char[length5];

	is5.read(buffer5, length5);
	intrusive_ptr<Stream> stream = new Spire::Common::Stream((unsigned char*)buffer, length);
	attachment->SetData((unsigned char*)buffer5, length5);
	attachment->SetDescription(L"Page footer picture of demo.");
	attachment->SetMimeType(L"image/png");
	doc->GetAttachments()->Add(attachment);
	float x = 50;
	intrusive_ptr<PdfTrueTypeFont> font2 = new PdfTrueTypeFont(L"Arial", 14.f, PdfFontStyle::Bold, true);
	intrusive_ptr<PointF> location = new PointF(x, y);

	wstring inputFile_F = input_path + L"Footer.png";
	ifstream is1(inputFile_F.c_str(), ifstream::in | ios::binary);

	is1.seekg(0, is1.end);
	int length1 = is.tellg();
	is1.seekg(0, is1.beg);

	char* buffer1 = new  char[length1];

	is1.read(buffer1, length1);
	intrusive_ptr<Stream> streamD = new Spire::Common::Stream((unsigned char*)buffer1, length1);

	wstring label = L"Sales Report Chart";
	intrusive_ptr<SizeF> size = font2->MeasureString(label.c_str());
	intrusive_ptr<RectangleF> bounds = new RectangleF(location, size);
	page->GetCanvas()->DrawString(label.c_str(), font2, PdfBrushes::GetDarkOrange(), bounds);
	bounds = new RectangleF(bounds->GetRight() + 3, bounds->GetTop(), font2->GetHeight() / 2, font2->GetHeight());

	//Create a PdfAttachmentAnnotation 
	LPCWSTR_S filename = L"SalesReportChart.png";
	intrusive_ptr<PdfAttachmentAnnotation> annotation1 = new PdfAttachmentAnnotation(bounds, filename, streamD);
	annotation1->SetColor(new PdfRGBColor(Spire::Common::Color::GetDarkOrange()));
	annotation1->SetFlags(PdfAnnotationFlags::ReadOnly);
	annotation1->SetIcon(PdfAttachmentIcon::Graph);
	annotation1->SetText(L"Sales Report Chart");
	//Add the annotation1
	page->GetAnnotationsWidget()->Add(annotation1);
	y = y + size->GetHeight() + 3;

	location = new PointF(x, y);
	label = L"Science Personification Boston";
	wstring inputFile_S = input_path + L"SciencePersonificationBoston.jpg";
	ifstream is2(inputFile_S.c_str(), ifstream::in | ios::binary);

	is2.seekg(0, is2.end);
	int length2 = is2.tellg();
	is2.seekg(0, is2.beg);

	char* buffer2 = new  char[length2];

	is2.read(buffer2, length2);
	intrusive_ptr<Stream> streamS = new Spire::Common::Stream((unsigned char*)buffer2, length2);
	size = font2->MeasureString(label.c_str());
	bounds = new RectangleF(location, size);
	page->GetCanvas()->DrawString(label.c_str(), font2, PdfBrushes::GetDarkOrange(), bounds);

	bounds = new RectangleF(bounds->GetRight() + 3, bounds->GetTop(), font2->GetHeight() / 2, font2->GetHeight());

	intrusive_ptr<PdfAttachmentAnnotation> annotation2 = new PdfAttachmentAnnotation(bounds, L"SciencePersonificationBoston.jpg", streamS);
	annotation2->SetColor(new PdfRGBColor(Spire::Common::Color::GetOrange()));
	annotation2->SetFlags(PdfAnnotationFlags::NoZoom);
	annotation2->SetIcon(PdfAttachmentIcon::PushPin);
	annotation2->SetText(L"SciencePersonificationBoston.jpg, from Wikipedia, the free encyclopedia");
	page->GetAnnotationsWidget()->Add(annotation2);
	y = y + size->GetHeight() + 2;

	location = new PointF(x, y);
	label = L"Picture of Science";
	wstring inputFile_w = input_path + L"SciencePersonificationBoston.jpg";
	ifstream is3(inputFile_w.c_str(), ifstream::in | ios::binary);

	is3.seekg(0, is3.end);
	int length3 = is3.tellg();
	is3.seekg(0, is3.beg);

	char* buffer3 = new  char[length3];

	is3.read(buffer3, length3);
	intrusive_ptr<Stream> streamW = new Spire::Common::Stream((unsigned char*)buffer3, length3);
	size = font2->MeasureString(label.c_str());
	bounds = new RectangleF(location, size);
	page->GetCanvas()->DrawString(label.c_str(), font2, PdfBrushes::GetDarkOrange(), bounds);
	bounds = new RectangleF(bounds->GetRight() + 3, bounds->GetTop(), font2->GetHeight() / 2, font2->GetHeight());
	intrusive_ptr<PdfAttachmentAnnotation> annotation3 = new PdfAttachmentAnnotation(bounds, L"Wikipedia_Science.png", streamW);
	annotation3->SetColor(new PdfRGBColor(Spire::Common::Color::GetSaddleBrown()));
	annotation3->SetFlags(PdfAnnotationFlags::Locked);
	annotation3->SetIcon(PdfAttachmentIcon::Tag);
	annotation3->SetText(L"Wikipedia_Science.png, from Wikipedia, the free encyclopedia");
	page->GetAnnotationsWidget()->Add(annotation3);
	y = y + size->GetHeight() + 2;

	location = new PointF(x, y);
	label = L"Hawaii Killer Font";
	wstring inputFile_tiff = input_path + L"PT_Serif-Caption-Web-Regular.ttf";
	ifstream is4(inputFile_tiff.c_str(), ifstream::in | ios::binary);

	is4.seekg(0, is4.end);
	int length4 = is4.tellg();
	is4.seekg(0, is4.beg);

	char* buffer4 = new  char[length4];

	is4.read(buffer4, length4);
	intrusive_ptr<Stream> streamP = new Spire::Common::Stream((unsigned char*)buffer4, length4);
	size = font2->MeasureString(label.c_str());
	bounds = new RectangleF(location, size);
	page->GetCanvas()->DrawString(label.c_str(), font2, PdfBrushes::GetDarkOrange(), bounds);
	bounds = new RectangleF(bounds->GetRight() + 3, bounds->GetTop(), font2->GetHeight() / 2, font2->GetHeight());
	intrusive_ptr<PdfAttachmentAnnotation> annotation4 = new PdfAttachmentAnnotation(bounds, L"PT_Serif-Caption-Web-Regular.ttf", streamP);
	annotation4->SetColor(new PdfRGBColor(Spire::Common::Color::GetCadetBlue()));
	annotation4->SetFlags(PdfAnnotationFlags::NoRotate);
	annotation4->SetIcon(PdfAttachmentIcon::Paperclip);
	annotation4->SetText(L"PT_Serif-Caption-Web-Regular, from https://company.paratype.com");
	page->GetAnnotationsWidget()->Add(annotation4);
	y = y + size->GetHeight() + 2;

	//Save pdf file
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}

