   M S B A M L   `   `   `  ����   GeocodeThru    GeocodeThru.MainWindow   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     -     � Icon$  images/earth globe.ico�� � Title$ GeocodeThru?��$	��545��$	��770��$	��770��$	��770��$	��545��$	��545�� � WindowStartupLocation$ CenterScreen=�.��� $��0q� � RowDefinitions �� $��30 ��� $��45 ��� $	��180 ��� $	��200 ��� $
��Auto � � ColumnDefinitions �� $	��392 ��� $
��Auto �$	��371��.+�]� $	��743��$��30��$
��Left=�$	��Top=�$��2��.+�o� $	��Top=�$��30��$
��Left=�$	��763��$��0,0,-19.8,0q�����  �� EndPoint$ 0.5,1&� � MappingMode$ RelativeToBoundingBox=� �� 
StartPoint$ 0.5,0&�.���  � Color$ 	#FFF1F1F1��	 � Offset$
	 0.25Q�� $ 	#FFAEAEAE��� $ 	#FFE4E4E4��$		 0.8Q�� $ 	#FFC9C9C9��$	 1Q�.@�m� -   
 m�Name$
 MenuOpen��$��_Open��$��14(�$��3,3,0,0q�.@�m� -   $
 MenuOpenXlsx��$
��XLSX�� m� Icon �� $��/images/excel.png�� �� BitmapScalingMode$ HighQuality=� �� EdgeMode$ Aliased=� � UseLayoutRounding . $��16��$��16��m� -   $
 MenuOpenAccdb��$��ACCDB��{�.   �� $��/images/access.png��$ HighQuality=�$ Aliased=� . $��16��$��16��m� -   $
 
MItemStart��$��_Start geocoding��$��	120,3,0,0q�$
��Bold"�$��14(�{�.  m� -   $
 
MItemReset��$��
_RESET APP��$��	120,3,0,0q�$��13(�m� -   $
 	MItemQuit��$��_Quit��$��	164,3,0,0q�$��13(�m� -   $
 
MItemAbout��$��_About��$��5,3,0,0q�$��13(��� -	    ��Name$ RtbProgress��$��1��$��2��$��0,0,0,10q� �� 
IsReadOnly . " �� VerticalScrollBarVisibility$ Visible=�$��Visible=�$	��360��$��Bottom=�$��Stretch=�$��0��$	��360��$
��Left=�$��3���� -
    ��Name$ 	GrbColMap��$��Column mapping��$��3��$��0��$��5,0,5,5q�{�.  .���� $��5q�.+��� $��0q� �� Orientation$ 
Horizontal=�$��Center=�.+��� $��5q�$��Ulice / Street��$	��200��$��RightToLeft=��� -    ��Name$ CbxUlice��$��5q�$	��150���� $��0q�$ 
Horizontal=�$��Center=�.+��� $��5q�$-��'Číslo popisné / Land registry number��$	��200��$��RightToLeft=��� -   $ CbxCp��$��5q�$	��150���� $��0q�$ 
Horizontal=�$��Center=�.+��� $��5q�$)��#Číslo orientační / House number��$	��200��$��RightToLeft=��� -   $ CbxCo��$��5q�$	��150���� $��0q�$ 
Horizontal=�$��Center=�.+��� $��5q�$��Obec / City��$	��200��$��RightToLeft=��� -   $ CbxObec��$��5q�$	��150���� $��0q�$ 
Horizontal=�$��Center=�.+��� $��5q�$��PSČ / Postal code��$	��200��$��RightToLeft=��� -   $ CbxPsc��$��5q�$	��150���� -   $ GrbProgress��$��Progress...��$��4��$��0��$��5,0,5,10q�.���� .+�� -    �Name$ ProgressBar��$��20��$��0,5,0,0q�$	��Top=�$��Center=�$	��350���� -   $ GrbApis��$
��APIs��$��0q�$��5,31,0,0q�$	��Top=�$��45��$��2��$��0��{�.  $��0��$��2��$
��Left=�$	��747��.���� $ 
Horizontal=�$��0,5,0,0q�.+��� -    ��Name$ CbxRuain��$��RUIAN��$
��Left=�$��	10,0,20,0q�$	��Top=��� -   $ CbxGm��$��GoogleMapsV3��$
��Left=�$��0,0,20,0q�$	��Top=�{�.  �� -   $ CbxHm��$��HereMaps��$
��Left=�$��0,0,20,0q�$	��Top=�{�.  �� -   $ CbxMq��$��OpenMapQuest��$
��Left=�$��0,0,20,0q�$	��Top=�{�.  �� -   $ CbxMcz��$��MapyCz��$
��Left=�$��0,0,20,0q�$	��Top=��� -   $ CbxOsm��$��OSMNominatim��$
��Left=�$��0,0,20,0q�$	��Top=��� -   $ CbxBm��$��BingMaps��$
��Left=�$��0q�$	��Top=�{�.  �� $��2��$��1��$	��Top=�$ 
Horizontal=�.+��� .+��� -   $ GrbWgsOrSjtsk��$��WGS or S-JSTK��$��5,0,5,5q�{�.  .���� .+��� -   $ CbxWgs��$	��WGS��$��2q� d� 	IsChecked$
 TrueY��� -   $ CbxSjtsk��$��S-JTSK��$��2q�$
 TrueY��� $��25,0,0,0q�.+��� $'��!From which row you want to start?��$��16��$	��183��$��5,5,5,0q��� $ 
Horizontal=�$��Center=�.+��� -    ��Name$ TxtbxFromRow��$��70��$��20��{�.  $��1�� �� TextAlignment$ Right=�$��5,0,5,5q��� -    ��Name$ TxtInterval��$��[1 - ?]��$��5,2,5,5q��� -    ��Name$ RectMethods��$��2��$��0��$
��Left=�$��34��$��	10,10,0,0q�$��1��$	��Top=�$	��743�� J� Opacity$	 0.5Q�$y�Hidden=� J� Effect -  'System.Windows.Media.Effects.BlurEffect     Radius$ 15Q�   
KernelType$	 Box=�����   �� Color$	  Red��! �� Opacity$	! 0.5Q��� -    $ RectDropdowns��$
��Left=�$	��200��$��	214,0,0,0q�$��3��$	��Top=�$	��173��$��0��$	 0.5Q�" J� RenderTransformOrigin$" 
0.476,0.48&�$y�Hidden=�   $ 15Q�$	 Box=����� $	  Red��$	! 0.5Q��� -!   $ 
RectCoords��$
��Left=�$��68��$��2��$	��Top=�$	��107��$��1��$	 0.5Q�$" 
0.476,0.48&�$y�Hidden=�   $ 15Q�$	 Box=����� $	  Red��$	! 0.5Q��� -"   $ GrbKeys��$��2��$��0��$��Keys for APIs��$��5,0,5,5q�{�.  .���� .+��� $ 
Horizontal=�.+��� .+��� $ 
Horizontal=�$��Center=�.+��� $��HereMaps - APP ID��$��5,6,0,5q�$��Right=�$��10(�$	��105���� -#   $ TxtboxHmAppId��$��5q�$��O5Xux7fAgmj4kSi67XbA��$	��235���� $ 
Horizontal=�$��Center=�.+��� $��HereMaps - APP CODE��$��5,2,0,5q�$��Right=�$��10(�$	��105���� -$   $ TxtboxHmAppCode��$��5,0,5,5q�$��cbeXpCRSVWo9kc17HbtHEA��$	��235���� -%   # ��Name$# ImgHm��$��15��$��15��$��images/question.png���� $��0q�$ 
Horizontal=�$��Center=�.+��� $��OpenMapQuest - KEY��$��5,6,0,5q�$��Right=�$��10(�$	��105���� -&   $ TxtboxMqKey��$��5q�$&�� 2i74YRgMWpE5GJoOlkFpy57yINNjQQ1V��$	��235���� -'   $# ImgMq��$��15��$��15��$��images/question.png���� $��0q�$ 
Horizontal=�$��Center=�.+��� $��GoogleMaps - KEY��$��5,6,0,5q�$��Right=�$��10(�$	��105���� -(   $ TxtboxGmKey��$��5q�$-��'AIzaSyDW2q2gTHY9pWsexT2doIiZYeom5p6BHUQ��$	��235���� -)   $# ImgGm��$��15��$��15��$��images/question.png���� $��0q�$ 
Horizontal=�$��Center=�.+��� $��BingMaps - KEY��$��5,6,0,5q�$��Right=�$��10(�$	��105���� -*   $ TxtboxBmKey��$��5q�$r��lwrIA0ucuQwsQUxP6OAZP~GcpFgalVzrLfG6E-qgBnaQ~AhcLmtVfz7TlpdyF12sCtimnxcho0RXl_eW_FRJIYlDjaiGiq-a1lc2cZOxmKBIb��$	��235���� -+   $# ImgBm��$��15��$��15��$��images/question.png��