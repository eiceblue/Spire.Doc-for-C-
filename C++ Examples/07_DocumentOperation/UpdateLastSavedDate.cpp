#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"UpdateLastSavedDate.docx";

	Document* document = new Document();

	//Load the document from disk
	document->LoadFromFile(inputFile.c_str());

	//Update the last saved date
	document->GetBuiltinDocumentProperties()->SetLastSaveDate(LocalTimeToGreenwishTime(DateTime::GetNow()));
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

DateTime* LocalTimeToGreenwishTime(DateTime* lacalTime)
{
	TimeZone* localTimeZone = TimeZone::GetCurrentTimeZone();
	TimeSpan* timeSpan = localTimeZone->GetUtcOffset(lacalTime);
	DateTime* greenwishTime = Spire::Common::DateTime::op_Subtraction(lacalTime, timeSpan);
	return greenwishTime;
}
