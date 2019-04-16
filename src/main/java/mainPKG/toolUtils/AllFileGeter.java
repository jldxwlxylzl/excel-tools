/**
 *
 */
package mainPKG.toolUtils;

import java.io.File;
import java.io.FileFilter;
import java.io.FilenameFilter;
import java.util.ArrayList;

public class AllFileGeter {

	/** 存放Root下所有文件的List*/
	private static ArrayList<File> allFilesList = null;
	/** 匹配文件名的正则表达式*/
	private static String regExp = null;

	/**
	 * 取得所有文件List的入口方法(默认找到所有文件)
	 * @param rootPath	查找所有文件的起点目录
	 * @return
	 */
	public static ArrayList<File> getAllFiles(String rootPath) {
		File rootDir = new File( rootPath );
		if ( !rootDir.exists() ) {
			System.out.println( "路径不存在 ---> " + rootPath );
			return null;
		} else if ( rootDir.isDirectory() ) {
			regExp = null;
			allFilesList = new ArrayList<File>();
			addAllFiles( rootDir );
			return allFilesList;
		} else {
			return null;
		}
	}

	/**
	 * 取得所有文件List的入口方法(过滤出文件名匹配fileNameRegExp的文件)
	 * @param rootPath	查找所有文件的起点目录
	 * @param fileNameRegExp	用于匹配文件名的正则表达式
	 * @return
	 */
	public static ArrayList<File> getAllFiles(String rootPath, String fileNameRegExp) {
		File rootDir = new File( rootPath );
		if ( !rootDir.exists() ) {
			return null;
		} else if ( rootDir.isDirectory() ) {
			regExp = fileNameRegExp;
			allFilesList = new ArrayList<File>();
			addAllFiles2( rootDir );
			return allFilesList;
		} else {
			return null;
		}
	}

	private static FileFilter fileFilter = null;
	public static ArrayList<File> getAllFiles(String rootPath, FileFilter _fileFilter) {
		File rootDir = new File( rootPath );
		if ( !rootDir.exists() ) {
			return null;
		} else if ( rootDir.isDirectory() ) {
			fileFilter = _fileFilter;
			allFilesList = new ArrayList<File>();
			addAllFiles3( rootDir );
			return allFilesList;
		} else {
			return null;
		}
	}

	private static FilenameFilter filenameFilter = null;
	public static ArrayList<File> getAllFiles(String rootPath, FilenameFilter _filenameFilter) {
		File rootDir = new File( rootPath );
		if ( !rootDir.exists() ) {
			return null;
		} else if ( rootDir.isDirectory() ) {
			filenameFilter = _filenameFilter;
			allFilesList = new ArrayList<File>();
			addAllFiles4( rootDir );
			return allFilesList;
		} else {
			return null;
		}
	}

	/**
	 * 递归添加文件的方法
	 * @param currentRootDir 递归过程中传递的文件或目录
	 * @return
	 */
	private static void addAllFiles(File currentRootDir) {
		if ( currentRootDir.isDirectory() ) {
			String[] pathList = currentRootDir.list();
			for (String path : pathList) {
				addAllFiles( new File( currentRootDir.getAbsolutePath() + File.separator + path ) );
			}
		} else {
			allFilesList.add( currentRootDir );
		}
	}

	private static void addAllFiles2(File currentRootDir) {
		if ( currentRootDir.isDirectory() ) {
			String[] nameList = currentRootDir.list();
			for (String name : nameList) {
				addAllFiles2( new File( currentRootDir, name ) );
			}
		} else if ( currentRootDir.getName().matches( regExp ) ) {
			allFilesList.add( currentRootDir );
		}
	}

	private static void addAllFiles3(File currentRootDir) {
		if (currentRootDir.isFile()) {
			allFilesList.add(currentRootDir);
		} else if ( currentRootDir.isDirectory() ) {
			for (File subFile : currentRootDir.listFiles( fileFilter )) {
				addAllFiles3(subFile);
			}
		}
	}

	private static void addAllFiles4(File currentRootDir) {
		if (currentRootDir.isFile()) {
			allFilesList.add(currentRootDir);
		} else if ( currentRootDir.isDirectory() ) {
			for (File subFile : currentRootDir.listFiles( filenameFilter )) {
				addAllFiles4(subFile);
			}
		}
	}

}
