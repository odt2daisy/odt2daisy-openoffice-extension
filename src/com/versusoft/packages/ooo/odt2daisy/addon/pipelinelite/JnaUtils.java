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

import com.sun.jna.Library;
import com.sun.jna.Native;


public class JnaUtils {

    private static CLibrary libc = (CLibrary) Native.loadLibrary("c", CLibrary.class);

    /**
     * Change mode (not for Windows)
     *
     */
    public static int chmod(String path, int mode){
        return libc.chmod(path, mode);
    }
}

interface CLibrary extends Library {
    public int chmod(String path, int mode);
}
