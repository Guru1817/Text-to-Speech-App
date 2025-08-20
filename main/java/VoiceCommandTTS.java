import com.google.cloud.texttospeech.v1.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;


public class VoiceCommandTTS {

    public static void main(String[] args) throws Exception {
        String excelFilePath = "commands.xlsx";  // put your Excel file here
        processExcel(excelFilePath);
    }

    // 🔹 Read Excel and generate speech for each row
    public static void processExcel(String filePath) throws Exception {
        FileInputStream fis = new FileInputStream(new File(filePath));
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        int rowCount = 0;
        for (Row row : sheet) {
            Cell cell = row.getCell(0);  // take text from first column
            if (cell != null) {
                String text = cell.getStringCellValue().trim();
                if (!text.isEmpty()) {
                    rowCount++;
                    String fileName = "output_" + rowCount + ".mp3";
                    synthesize(text, fileName);
                }
            }
        }
        workbook.close();
        fis.close();
    }

    // 🔹 Text-to-Speech function
    public static void synthesize(String text, String fileName) throws Exception {
        try (TextToSpeechClient textToSpeechClient = TextToSpeechClient.create()) {
            SynthesisInput input = SynthesisInput.newBuilder()
                    .setText(text)
                    .build();

            VoiceSelectionParams voice = VoiceSelectionParams.newBuilder()
                    .setLanguageCode("en-IN") // Fallback to a supported language
                    .setName("en-IN-Wavenet-A") // Specify a valid voice
                    .setSsmlGender(SsmlVoiceGender.NEUTRAL)
                    .build();

            AudioConfig audioConfig = AudioConfig.newBuilder()
                    .setAudioEncoding(AudioEncoding.MP3)
                    .build();

            try {
                SynthesizeSpeechResponse response =
                        textToSpeechClient.synthesizeSpeech(input, voice, audioConfig);
                try (FileOutputStream out = new FileOutputStream(fileName)) {
                    out.write(response.getAudioContent().toByteArray());
                    System.out.println("✅ Generated: " + fileName + " for text: " + text);
                }
            } catch (Exception e) {
                System.err.println("Error synthesizing speech for text: " + text);
                e.printStackTrace();
            }
        } catch (Exception e) {
            System.err.println("Error initializing TextToSpeechClient: " + e.getMessage());
            e.printStackTrace();
        }
    }
}