#include "pch.h"
using namespace Spire::Doc;


intrusive_ptr<DateTime>LocalTimeToGreenwishTime(intrusive_ptr<DateTime> lacalTime);
int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"UpdateLastSavedDate.docx";

	intrusive_ptr<Document> document = new Document();

				//Load the document from disk
				document->LoadFromFile(inputFile.c_str());

				//Update the last saved date
				document->GetBuiltinDocumentProperties()->SetLastSaveDate(LocalTimeToGreenwishTime(DateTime::GetNow()));
				document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
				document->Close();
}

intrusive_ptr<DateTime> LocalTimeToGreenwishTime(intrusive_ptr<DateTime> lacalTime)
{
	intrusive_ptr<TimeZone> localTimeZone = TimeZone::GetCurrentTimeZone();
	intrusive_ptr<TimeSpan> timeSpan = localTimeZone->GetUtcOffset(lacalTime);
	//DateTime greenwishTime = lacalTime - timeSpan;
	intrusive_ptr<DateTime> greenwishTime = DateTime::op_Subtraction(lacalTime, timeSpan);
	return greenwishTime;
}
