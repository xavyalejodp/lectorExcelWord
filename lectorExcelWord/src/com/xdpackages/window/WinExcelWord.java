package com.xdpackages.window;

import java.awt.EventQueue;
import java.io.File;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileSystemView;

import javax.swing.JFrame;
import javax.swing.JButton;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;


import javax.swing.JLabel;
import javax.swing.JTextField;


public class WinExcelWord {

	private JFrame frame;
	private JTextField txtArchivoExcel;	
	private JTextField txtArchivoWord;
	private String nombreArchivoExcel;
	private String nombreArchivoWord;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					WinExcelWord window = new WinExcelWord();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public WinExcelWord() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.setBounds(100, 100, 516, 220);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		
		JLabel lblSeleccioneElArchivo = new JLabel("Seleccione el archivo Excel:");
		lblSeleccioneElArchivo.setBounds(10, 25, 162, 14);
		frame.getContentPane().add(lblSeleccioneElArchivo);
		
		txtArchivoExcel = new JTextField();
		txtArchivoExcel.setBounds(182, 22, 175, 20);
		frame.getContentPane().add(txtArchivoExcel);
		txtArchivoExcel.setColumns(10);
		
		JButton btnBuscar = new JButton("Buscar");
		btnBuscar.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());

				int returnValue = jfc.showOpenDialog(null);
				// int returnValue = jfc.showSaveDialog(null);

				if (returnValue == JFileChooser.APPROVE_OPTION) {
					File selectedFile = jfc.getSelectedFile();
					System.out.println(selectedFile.getAbsolutePath());
					nombreArchivoExcel =selectedFile.getAbsolutePath();
					txtArchivoExcel.setText(nombreArchivoExcel);
				}
			}
		});
		btnBuscar.setBounds(367, 21, 123, 23);
		frame.getContentPane().add(btnBuscar);
		
		JButton btnProcesarArchivo = new JButton("Procesar Archivo");
		btnProcesarArchivo.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
			}
		});
		btnProcesarArchivo.setBounds(160, 127, 175, 23);
		frame.getContentPane().add(btnProcesarArchivo);
		
		JLabel lblNombreDeArchivo = new JLabel("Nombre de archivo a guardar:");
		lblNombreDeArchivo.setBounds(10, 67, 153, 14);
		frame.getContentPane().add(lblNombreDeArchivo);
		
		txtArchivoWord = new JTextField();
		txtArchivoWord.setBounds(182, 64, 175, 20);
		frame.getContentPane().add(txtArchivoWord);
		txtArchivoWord.setColumns(10);
		
		JButton btnDirectorioGuardar = new JButton("Directorio Guardar");
		btnDirectorioGuardar.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());

				int returnValue = jfc.showOpenDialog(null);
				// int returnValue = jfc.showSaveDialog(null);

				if (returnValue == JFileChooser.APPROVE_OPTION) {
					File selectedFile = jfc.getSelectedFile();
					System.out.println(selectedFile.getAbsolutePath());
					nombreArchivoWord =selectedFile.getAbsolutePath();
					txtArchivoWord.setText(nombreArchivoWord);
				}
			}
		});
		btnDirectorioGuardar.setBounds(367, 63, 123, 23);
		frame.getContentPane().add(btnDirectorioGuardar);
	}
}
