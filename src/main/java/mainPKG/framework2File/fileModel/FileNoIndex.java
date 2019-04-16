package mainPKG.framework2File.fileModel;

public class FileNoIndex extends AbsSrcFile<BasicSrcLine> {

	public FileNoIndex(String _absolutePath) {
		super(_absolutePath);
	}

	@Override
	public void readFile() throws Exception {
		readFile(null, BasicSrcLine.class);
	}

	@Override
	public void readFile(String encoding) throws Exception {
		readFile(encoding, BasicSrcLine.class);
	}

}
