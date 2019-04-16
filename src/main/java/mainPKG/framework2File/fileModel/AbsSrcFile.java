package mainPKG.framework2File.fileModel;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;

public abstract class AbsSrcFile<Line extends BasicSrcLine> {
	protected Class<Line> CLASS_LINE;

	protected String absolutePath = null;

	public String getAbsolutePath() {
		return absolutePath;
	}

	public AbsSrcFile(String _absolutePath) {
		absolutePath = _absolutePath;
	}

	public Line firstLine = null;
	public Line lastLine = null;

	protected List<Line> allLines = new ArrayList<Line>();

	public List<Line> getAllLines() {
		return allLines;
	}

	public Line lineAt(int lineIndex) {
		return allLines.get(lineIndex);
	}

	public int lineCnt() {
		return allLines.size();
	}

	protected void readFile(String encoding, Class<Line> classLine) throws Exception {
		CLASS_LINE = classLine;

		firstLine = null;
		lastLine = null;
		allLines.clear();

		FileInputStream fis = new FileInputStream(absolutePath);
		InputStreamReader isr;
		if (encoding == null) {
			isr = new InputStreamReader(fis);
		} else {
			isr = new InputStreamReader(fis, encoding);
		}
		BufferedReader br = new BufferedReader(isr);

		Line lineTmp = null;
		String txt = br.readLine();
		if (txt != null) {
			lineTmp = creatALine(txt);
			firstLine = lineTmp;
			lastLine = lineTmp;
			allLines.add(firstLine);
		}

		while ((txt = br.readLine()) != null) {
			lineTmp = creatALine(lineTmp);
			lineTmp.setSrcContent(txt);

			lastLine = lineTmp;
			allLines.add(lineTmp);
		}

		br.close();
	}

	public abstract void readFile() throws Exception;
	
	public abstract void readFile(String encoding) throws Exception;

	public void writeOut(String outputPath) throws IOException {
		FileWriter fileWriter = new FileWriter(outputPath);

		List<String> outputList = makeOutputList();

		for (String outLine : outputList) {
			fileWriter.write(outLine);
			fileWriter.write(System.lineSeparator());
		}
		fileWriter.close();
	}

	protected List<String> makeOutputList() {
		ArrayList<String> outList = new ArrayList<String>();
		for (Line basicSrcLine : allLines) {
			outList.add(basicSrcLine.getSrcContent());
		}
		return outList;
	}

	@Override
	public String toString() {
		return "[File ::-> " + absolutePath + "]";
	}

	private Line creatALine(String txt) throws Exception {
		Line rtnLine = CLASS_LINE.getConstructor(String.class).newInstance(txt);
		return rtnLine;
	}

	private Line creatALine(Line line) throws Exception {
		Line rtnLine = CLASS_LINE.getConstructor(CLASS_LINE).newInstance(line);
		return rtnLine;
	}
}
