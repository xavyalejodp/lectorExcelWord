package com.xdpackages.config;

import java.util.MissingResourceException;
import java.util.ResourceBundle;

public class Configuration {
	
	/** The Constant BUNDLE_NAME. */
	private static final String BUNDLE_NAME = "com.xdpackages.config.configuration"; //$NON-NLS-1$

	/** The Constant RESOURCE_BUNDLE. */
	private static final ResourceBundle RESOURCE_BUNDLE = ResourceBundle
			.getBundle(BUNDLE_NAME);
    
    /** The Constant SEPARADOR. */
    private static final String SEPARADOR = "&param";

    /**
     * Instantiates a new application.
     */
	private Configuration() {
	}

	/**
	 * Gets the string.
	 * Recupera la propiedad referente al key pasado
	 * 
	 * @param key the key
	 * 
	 * @return the string
	 */
	public static String getString(String key) {
		try {
			return RESOURCE_BUNDLE.getString(key);
		} catch (MissingResourceException e) {
			return '!' + key + '!';
		}
	}

    /**
     * Gets the string.
     * Recupera la propiedad referente al key pasado complmentada
     * su perzonalización con los parámetros adicionales enviados.
     * 
     * @param key the key
     * @param parameters the parameters
     * 
     * @return the string
     */
    public static String getString(String key, String... parameters) {
        try {
            String message = "";
            String[] partsMessage = RESOURCE_BUNDLE.getString(key)
                                                   .split(SEPARADOR);

            for (String part : partsMessage) {
                try {
                    int indice = Integer.parseInt(part);

                    if (parameters.length > indice) {
                        message = message.concat(parameters[indice]);
                    }
                } catch (NumberFormatException e) {
                    message = message.concat(part);
                }
            }

            return message;
        } catch (MissingResourceException e) {
            return '!' + key + '!';
        }
    }

}
