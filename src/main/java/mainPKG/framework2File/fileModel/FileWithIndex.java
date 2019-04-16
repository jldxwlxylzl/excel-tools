package mainPKG.framework2File.fileModel;

public class FileWithIndex extends AbsSrcFile<SrcIndexLine> {

	public FileWithIndex(String _absolutePath) {
		super(_absolutePath);
	}

	@Override
	public void readFile() throws Exception {
		readFile(null, SrcIndexLine.class);
		firstLine.setLineIndex(0, true);
	}

	@Override
	public void readFile(String encoding) throws Exception {
		readFile(encoding, SrcIndexLine.class);
		firstLine.setLineIndex(0, true);
	}
	
}
