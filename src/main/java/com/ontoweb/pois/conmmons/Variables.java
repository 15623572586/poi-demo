package com.ontoweb.pois.conmmons;

import lombok.Data;

import java.util.HashMap;

@Data
public class Variables {

    private static final HashMap<String, String> DATA_MAPPING = new HashMap<String, String>(){
        {
            put("N", "SIGNAL_UNIQUE_CODE"); // source, target
            put("F", "SIGNAL_NAME"); // source, target
            put("D", "TESTED_EQPT_NINECODE"); // source, target
        }
    };

}
