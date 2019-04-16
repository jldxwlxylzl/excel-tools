package mainPKG.framework2File.taskModel;

import java.io.File;
import java.io.FileFilter;
import java.io.FileNotFoundException;
import java.io.FilenameFilter;
import java.util.ArrayList;
import java.util.List;

public class WorkManager {
	
	private File rootDir = null;
	public File getRootDir() {
		return rootDir;
	}
	public void setRootDir(File rootDir) {
		this.rootDir = rootDir;
	}
	public void setRootDir(String rootDir) {
		setRootDir( new File( rootDir ) );
	}

	public WorkManager(String _rootDir) {
		rootDir = new File( _rootDir );
		if ( !rootDir.exists() ) {
			try {
				throw new FileNotFoundException( "RootDir Not Found ---> [" + _rootDir + "]" );
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			}
		}
	}
	private FilenameFilter filter = null;
	public void setFileFilter(boolean containsSubDir, String... _suffixs) {
		if ( _suffixs==null || _suffixs.length==0 ) {
			filter = null;
			return;
		}
		filter = new FilenameFilter() {
			@Override
			public boolean accept(File dir, String name) {
				if (dir.equals(rootDir)) {
					if (new File(dir, name).isDirectory()) {
						return containsSubDir;
					}
					for (String suffix : _suffixs) {
						if ( name.endsWith( suffix ) ) {
							return true;
						}
					}
				} else {
					if (new File(dir, name).isDirectory()) {
						return true;
					}
					for (String suffix : _suffixs) {
						if ( name.endsWith( suffix ) ) {
							return true;
						}
					}
				}
				return false;
			}
		};
	}
	public void setFileFilter(FilenameFilter _filter) {
		filter = _filter;
	}
	public void setFileFilter(FileFilter _filter) {
		filter = new FilenameFilter() {
			@Override
			public boolean accept(File dir, String name) {
				return _filter.accept( new File(dir, name) );
			}
		};
	}
	private List<File> fileContainer = null;
	private void listAllFiles() {
		fileContainer = new ArrayList<File>();
		if ( filter == null ) {
			findAllFiles(rootDir);
		} else {
			findAllFilesWithFilter(rootDir);
		}
	}
	private void findAllFiles(File subDir) {
		if ( subDir.isFile() ) {
			fileContainer.add( subDir );
		} else {
			File[] listFiles = subDir.listFiles();
			for (File file : listFiles) {
				findAllFiles( file );
			}
		}
	}
	private void findAllFilesWithFilter(File subDir) {
		if ( subDir.isFile() ) {
			fileContainer.add( subDir );
		} else {
			File[] listFiles = subDir.listFiles( filter );
			for (File file : listFiles) {
				findAllFilesWithFilter( file );
			}
		}
	}
	
	public void visitAllFiles(WorkFileTask _fileTask) {
		listAllFiles();
		int allCnt = fileContainer.size();
		for (int i = 0; i < allCnt; i++) {
			File fileCurrent = fileContainer.get( i );
			_fileTask.task_Init( fileCurrent, i, allCnt );
			_fileTask.task_ToDo( fileCurrent, i, allCnt );
			_fileTask.task_End( fileCurrent, i, allCnt );
		}
	}
	
}
