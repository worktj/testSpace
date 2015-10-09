package gui.hilfsklassen;

import java.sql.SQLException;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Calendar;


import com.sun.glass.ui.Screen;
import com.sun.star.awt.PosSize;
import com.sun.star.beans.PropertyValue;
import com.sun.star.bridge.UnoUrlResolver;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.connection.ConnectionSetupException;
import com.sun.star.connection.NoConnectException;
import com.sun.star.document.XDocumentInsertable;
import com.sun.star.document.XEventBroadcaster;
import com.sun.star.frame.XController;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XModel;
import com.sun.star.io.IOException;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XEventListener;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.text.XTextViewCursor;
import com.sun.star.text.XTextViewCursorSupplier;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XInterface;
import com.sun.star.util.XReplaceDescriptor;
import com.sun.star.util.XReplaceable;
import com.sun.star.util.XSearchDescriptor;

import javafx.concurrent.Task;
import testtools.S;

public class OpenOfficeUNO {

	private XComponent xComponent;

	public OpenOfficeUNO(long _empfaenger_id) {
	}

	/**
	 * SERVICE maybe? to return and be able to close wrong opened template!
	 *
	 * @param url Pfad zur Vorlage Datei
	 * @param newName neuer Name für diese Vorlage Datei
	 *
	 */
	public void openDokumentVorlageZurBearbeitung(String url, String newName) {

		Task<Void> t = new Task<Void>() {

			@Override
			protected Void call() throws Exception {
				com.sun.star.uno.XComponentContext xContext = null;

				xContext = Bootstrap.createInitialComponentContext(null);

				XUnoUrlResolver urlResolver = UnoUrlResolver.create(xContext);

				Object initialObject = urlResolver
						.resolve("uno:socket,host=localhost,port=8100;urp;StarOffice.ServiceManager");
							   // uno:socket,host=localhost,port=2083;urp;StarOffice.ServiceManager

				XMultiComponentFactory xMCF = (XMultiComponentFactory) UnoRuntime
						.queryInterface(XMultiComponentFactory.class,
								initialObject);
				try {
					// get the remote office service manager
					Object oDesktop = xMCF.createInstanceWithContext(
							"com.sun.star.frame.Desktop", xContext);

//					XDesktop xDesktop = (XDesktop) UnoRuntime.queryInterface(com.sun.star.frame.XDesktop.class, oDesktop);

					com.sun.star.frame.XComponentLoader xCLoader = UnoRuntime
							.queryInterface(
									com.sun.star.frame.XComponentLoader.class,
									oDesktop);

					PropertyValue[] loadProps = new PropertyValue[2];
					loadProps[0] = new PropertyValue();
					loadProps[0].Name = "AsTemplate";
					loadProps[0].Value = new Boolean(true);
					loadProps[1] = new PropertyValue();
					loadProps[1].Name = "Overwrite";
					loadProps[1].Value = new Boolean(true);

					xComponent = xCLoader.loadComponentFromURL("file:///" + url,
							"_blank", 0, loadProps);

					XTextDocument xtextdoc = UnoRuntime
							.queryInterface(XTextDocument.class,
									xComponent);

//					while(insertReplacements(xtextdoc, inhaltID.getAktenID(), beteiligung_empfaenger_id)) {
//
//					}

					com.sun.star.frame.XStorable xStorable =
			                (com.sun.star.frame.XStorable)UnoRuntime.queryInterface(
			                    com.sun.star.frame.XStorable.class, xComponent );
					// TODO erstelle einen aussagekräftigen Dateinamen @uniquename
					String uniquename ="C:/Users/Thomas.Jahn/Documents/OOtest/Documents/OOtest/" +newName+ Math.random() + ".odt";
//					C:\Users\Thomas.Jahn\Documents\OOtest
//					 System.getProperty("user.home")

					xStorable.storeAsURL("file:///" + url, new PropertyValue[0]);
//					xStorable.store();

//					ArrayList<SimpleTyp> dokumenttyp = new ArrayList<SimpleTyp>();
//					dokumenttyp.add(new SimpleTyp(5, "Sonstige"));

					S.o3("DokmentVorlage gespeichert unter :"+xStorable.getLocation());

//					DokumentDAO.insert(new ModelDokument(0L,
//							inhaltID.getAktenID(),
//							uniquename,
//							true,
//							new Timestamp(Calendar.getInstance().getTimeInMillis()),
//							90,
//							true,
//							dokumenttyp));

					com.sun.star.frame.XModel xModel = (com.sun.star.frame.XModel) UnoRuntime
							.queryInterface(com.sun.star.frame.XModel.class, xComponent);

					com.sun.star.frame.XFrame xFrame = xModel.getCurrentController().getFrame();

					XEventBroadcaster xdoceventbroadc = (XEventBroadcaster) UnoRuntime
							.queryInterface(XEventBroadcaster.class, xComponent);

					xdoceventbroadc.addEventListener(new com.sun.star.document.XEventListener() {

						@Override
						public void disposing(EventObject arg0) {
							// TODO Auto-generated method stub
							S.o3("DOCEVENT disposing " + arg0.toString() );

						}

						@Override
						public void notifyEvent(com.sun.star.document.EventObject arg0) {
							S.o3("DOCEVENT " + arg0.EventName );

							if(arg0.EventName.equals("OnSave") || arg0.EventName.equals("OnSaveAs")) {
								S.o3("SAVE" + uniquename);
							}

						}
					});

					Screen secondScreen = Screen.getScreens().get(1);

					int framex = secondScreen.getX() + 6;
					int framey = secondScreen.getY();
					int frameh = secondScreen.getVisibleHeight() - 20;
					int framew = secondScreen.getVisibleWidth() / 2;

					xFrame.getContainerWindow().setPosSize(framex, framey, framew, frameh, PosSize.POSSIZE);

//					GUIManager.resizeSecStageHalf();

				} catch (Exception e) {
					System.err.println(" Exception " + e);
					e.printStackTrace(System.err);
				}

				return null;

			}
		};

		t.run();
	}


	/**
	 *
	 * Extrahiere den Rohtext aus der Office Datei ohne die
	 *
	 * @param url Pfad zur OpenOffice Datei
	 * TODO prüfe ob url auf odt endet.
	 * @return
	 * @throws Exception
	 * @throws NoConnectException
	 * @throws ConnectionSetupException
	 * @throws Exception
	 * @throws IOException
	 */
	public String getVorschauStringRaw(String url)
			throws Exception, NoConnectException, ConnectionSetupException, com.sun.star.uno.Exception, IOException {
		com.sun.star.uno.XComponentContext _xContext = Bootstrap
				.createInitialComponentContext(null);

		XUnoUrlResolver urlResolver = UnoUrlResolver.create(_xContext);

		Object initialObject = urlResolver
				.resolve("uno:socket,host=localhost,port=8100;urp;StarOffice.ServiceManager");

		XMultiComponentFactory xMCF = (XMultiComponentFactory) UnoRuntime
				.queryInterface(XMultiComponentFactory.class, initialObject);

		Object oDesktop = xMCF.createInstanceWithContext(
				"com.sun.star.frame.Desktop", _xContext);

		com.sun.star.frame.XComponentLoader xCLoader = UnoRuntime
				.queryInterface(com.sun.star.frame.XComponentLoader.class,
						oDesktop);

		PropertyValue[] loadProps = new PropertyValue[1];
		loadProps[0] = new PropertyValue();
		loadProps[0].Name = "Hidden";
		loadProps[0].Value = new Boolean(true);

		XComponent xCurrentComponent = xCLoader.loadComponentFromURL("file:///"
				+ url, "_blank", 0, loadProps);

		XTextDocument xDoc = UnoRuntime.queryInterface(
				com.sun.star.text.XTextDocument.class, xCurrentComponent);

		String result = xDoc.getText().getString();

		com.sun.star.frame.XModel xModel = (com.sun.star.frame.XModel) UnoRuntime
				.queryInterface(com.sun.star.frame.XModel.class, xDoc);

		if (xModel != null) {
			com.sun.star.util.XCloseable xCloseable = (com.sun.star.util.XCloseable) UnoRuntime
					.queryInterface(com.sun.star.util.XCloseable.class, xModel);

			if (xCloseable != null) {
				try {
					xCloseable.close(true);
				} catch (com.sun.star.util.CloseVetoException exCloseVeto) {
				}
			} else {
				com.sun.star.lang.XComponent xDisposeable = (com.sun.star.lang.XComponent) UnoRuntime
						.queryInterface(com.sun.star.lang.XComponent.class,
								xModel);
				xDisposeable.dispose();
			}
		}


		return result;

	}

	//	public void setCloseListener(KontextTab kt) {
	public void setCloseListener() {
		com.sun.star.frame.XModel xModel = (com.sun.star.frame.XModel) UnoRuntime
				.queryInterface(com.sun.star.frame.XModel.class, xComponent);

		com.sun.star.frame.XFrame xFrame = xModel.getCurrentController().getFrame();

		xFrame.addEventListener(new XEventListener() {
			@Override
			public void disposing(EventObject arg0) {
//				InhaltManager.closeContextTab(inhaltID, kt);
			}
		});
	}

}
