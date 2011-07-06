/**
 *  odt2daisy - OpenDocument to DAISY XML/Audio
 *
 *  (c) Copyright 2008 - 2009 by Vincent Spiewak, All Rights Reserved.
 *
 *  This program is free software; you can redistribute it and/or modify
 *  it under the terms of the GNU General Lesser Public License as published by
 *  the Free Software Foundation; either version 3 of the License, or
 *  (at your option) any later version.
 *
 *  This program is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *  GNU Lesser General Public License for more details.
 *
 *  You should have received a copy of the GNU Lesser General Public License along
 *  with this program; if not, write to the Free Software Foundation, Inc.,
 *  51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
 */
package com.versusoft.packages.ooo.odt2daisy.addon.pipelinelite;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.logging.Logger;

/**
 *
 * @author Vincent SPIEWAK
 */
public class PipelineLite {

    public final static boolean IS_WINDOWS = System.getProperty("os.name").toLowerCase().contains("windows");
    public final static boolean IS_OSX = System.getProperty("os.name").toLowerCase().contains("mac os x");
    //
    public final static String USERMODE_FOLDER = System.getProperty("user.home") + System.getProperty("file.separator");
    public final static String TMPMODE_FOLDER = System.getProperty("java.io.tmpdir");
    //
    private final static String PIPELINE_EXEC_NAME = "pipeline-lite";
    private final static String PIPELINE_EXEC_EXT = getPipelineLiteExt();
    private final static String PIPELINE_EXEC = PIPELINE_EXEC_NAME + PIPELINE_EXEC_EXT;
    //
    public final static String PIPELINE_FOLDER_NAME = ".pipeline-lite" + System.getProperty("file.separator");
    private final static String EXT_BIN_FOLDER = "ext" + System.getProperty("file.separator");
    private final static String SCRIPT_FOLDER = "scripts" + System.getProperty("file.separator");
    private final static String NARRATOR_SCRIPT_NAME = "DAISY Pipeline TTS Narrator.taskScript";
    private final static String PIPELINE_ZIP_PATH = "/com/versusoft/packages/ooo/odt2daisy/addon/pipelinelite/build/pipeline-lite.zip";
    //
    private final static String PIPELINE_VERSION = "VERSION_20100325";
    //
    public final static long REQUIRED_KB = 60000;
    private final Logger logger;

    /**
     * 
     * @param logger
     */
    public PipelineLite(Logger logger) {
        this.logger = logger;
    }

    /**
     * Return extension of PipelineLite binary depending on the platform
     * @return extension with the dot
     */
    private static String getPipelineLiteExt() {
        if (IS_WINDOWS) {
            return ".exe";
        } else {
            return ".sh";
        }
    }

    /**
     * Check if PipelineLite is extracted
     *
     * @param folder: folder to check
     * @return true if extracted
     */
    public static boolean isExctracted(String folder) {
        return new File(folder + PIPELINE_FOLDER_NAME).canRead();
    }

    /**
     *
     * Check if PipelineLite is not outdated (using VERSION_* file)
     *
     * @param folder
     * @return
     */
    public static boolean isOutdated(String folder) {
        return !(new File(folder + PIPELINE_FOLDER_NAME + PIPELINE_VERSION).canRead());
    }

    /**
     *
     * Extract PipelineLite in the folder given in argument
     *
     * @param folder
     * @throws ExtractErrorException
     */
    public void extract(String folder) throws Exception {
        try {

            logger.info("DEBUG: extract to:" + folder);

            File tmpFile = File.createTempFile("temp", null);

            logger.info("DEBUG: tmp file: " + tmpFile);

            logger.info("DEBUG: copy outside");

            // Copy PipelineLite archive outside
            ZipUtils.copy(
                    PipelineLite.class.getResourceAsStream(PIPELINE_ZIP_PATH),
                    new FileOutputStream(tmpFile));

            // Remove Old PipelineLite
            File pipelineLiteFile = new File(folder + PIPELINE_FOLDER_NAME);
            if (pipelineLiteFile.exists()) {
                logger.info("DEBUG: delete old version");
                if (!deleteDir(pipelineLiteFile)) {
                    throw new Exception("Can't delete old PipelineLite folder: " + pipelineLiteFile);
                }
            }

            logger.info("DEBUG: extract to: " + pipelineLiteFile);

            // Extract Archive
            ZipUtils.unzipArchive(
                    tmpFile,
                    pipelineLiteFile);

            // Try to hide folder using attrib.exe
            if (IS_WINDOWS) {
                try {
                    String cmd[] = {"attrib", "+h", pipelineLiteFile.getAbsolutePath()};
                    Runtime.getRuntime().exec(cmd);
                } catch(Exception e){
                    // not a big deal if not hidden ...
                    // continue
                }
            }

        } catch (Exception e) {
            throw new Exception(e);
        }
    }

    /**
     *
     * Launch PipelineLite Narrator.
     * See also http://data.daisy.org/projects/pipeline/doc/scripts/Narrator-DtbookToDaisy.html 
     *
     * @param folder
     * @param fileIn Valid DTBook 2005 file.
     * @param fileOut The base directory of the output.
     * @param fixRoutine Determines whether Narrator should repair and tidy a suboptimal DTBook document. (See http://data.daisy.org/projects/pipeline/doc/scripts/DTBookFix.html .)
     * @param sentDetection Selects whether to apply sentence detection to the input document. This is required for synchronising highlighting with the synthetic speech in software-based DAISY players.
     * @param bitrate The bitrate of the generated MP3 files. A higher value will result in better sound quality but the audio files will be larger.
     * @return Return code from the Narrator. 0 = OK. 
     * @throws Exception
     */
    public int launchNarrator(String folder, String fileIn, String fileOut, boolean fixRoutine, boolean sentDetection, int bitrate) throws Exception {
        Runtime runtime = Runtime.getRuntime();
        Process process;

        String extracted_dir = folder;
        String pipeline_dir = extracted_dir + PIPELINE_FOLDER_NAME;
        String pipeline_exec = pipeline_dir + PIPELINE_EXEC;

        // chmod +x pipeline executable
        // (no need for win)
        if (!IS_WINDOWS) {

            logger.info("DEBUG: chmod on " + pipeline_exec);
            JnaUtils.chmod(pipeline_exec, 0755);

            // chmod +x binairies
            File ext_folder = new File(pipeline_dir + EXT_BIN_FOLDER);
            String fileList[] = ext_folder.list();

            for (int i = 0; i < fileList.length; i++) {
                logger.info("DEBUG: chmod on " + fileList[i]);
                JnaUtils.chmod(pipeline_dir + EXT_BIN_FOLDER + fileList[i], 0755);
            }
        }

        String fixRoutineParamValue;
        if(fixRoutine){
            fixRoutineParamValue = "REPAIR_TIDY_NARRATOR";
        } else {
            fixRoutineParamValue = "NOTHING";
        }

        String exec_cmd[] = {
            pipeline_exec,
            "-x", 
            //"-q",
            "-s",
            SCRIPT_FOLDER + NARRATOR_SCRIPT_NAME,
            "-p ",
            "input=" + fileIn,
            "outputPath=" + fileOut,
            "doSentDetection=" + sentDetection,
            "dtbookFix=" + fixRoutineParamValue,
            "bitrate=" + bitrate
        };

        String env[] = null;

        for (int i = 0; i < exec_cmd.length; i++) {
            logger.info("DEBUG: cmd:" + exec_cmd[i]);
        }

        logger.info("DEBUG: runtime.exec start");

        process = runtime.exec(exec_cmd, env, new File(pipeline_dir));
        InputStream stderr = process.getErrorStream();
        InputStreamReader isr = new InputStreamReader(stderr);
        BufferedReader br = new BufferedReader(isr);
        String line = null;
        String errors = "";
        while ((line = br.readLine()) != null) {
            errors += line + "\n";
            logger.info("DEBUG: " + errors);

        }

        return process.waitFor();
    }

    /**
     *
     * Delete a directory (empty or non-empty)
     *
     * @param dir
     * @return true on success, stop if an error occur
     */
    private static boolean deleteDir(File dir) {
        if (dir.isDirectory()) {
            String[] children = dir.list();
            for (int i = 0; i <
                    children.length; i++) {
                boolean success = deleteDir(new File(dir, children[i]));
                if (!success) {
                    return false;
                }

            }
        }

        // The directory is now empty so delete it
        return dir.delete();
    }
}
