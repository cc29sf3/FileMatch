   M S B A M L   `   `   `  ����H  DManual_Import, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullA%clr-namespace:Manual_Import.ConvertorManual_Import.Convertor  5   	       Manual_Import.SetUp   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   7local%clr-namespace:Manual_Import.Convertor     5   	     � Title$  提取页数设置��5   	   $	��200��6   $	��320��6%    � 
ResizeMode$ NoResize=�61    � WindowStartupLocation$ CenterScreen=�6G   .��5      �  � 	Resources 5   
   1  +Manual_Import.Convertor.ForegroundConvertor   	F_convert �� 
TargetType �� Property �� Value �� ElementName �� Path	 �� 	Converter��%  &	        '��      '��      7     5      6<   �� 5	      " ��6   .�5
      ��  � {� 6.   �� $ checkBox��$ 	IsChecked�6f   5      �� 5      " ��6   .�5      ��  � �� 6/   �� $ checkBox��$ 	IsChecked�8	   6�   5      5      .+�5   
   �� 
 �� Orientation$
 Vertical=�6   $��
0,10,108,0q�6,   .+�5      �� $
 
Horizontal=�6   $��
40,10,10,0q�62   .+�5      �� $��68��6   提取PDF前6/   �� 5      -   5        $��pdfFront��5   6   $��40��6   $��3,0,3,0q�6%   6H   �� 5      页6   5      �� 5      $
 
Horizontal=�6   $��
40,10,10,0q�62   .+�5      �� $��68��6   提取PDF后60   �� 5      -   5        $��pdfBack��5   6   $��40��6   $��3,0,3,0q�6%   6G   �� 5      页6   5      �� 5      $
 
Horizontal=�6   $��
40,10,10,0q�62   .+�5      �� 提取Word前6%   �� 5      -   5        $��	wordFront��5   6   $��40��6   $��3,0,3,0q�6%   6I   �� 5      页6   5      �� 5       $
 
Horizontal=�6   $��
40,10,10,0q�62   .+�5!      �� 提取Word后6%   �� 5"      -   5        $��wordBack��5"   6   $��40��6   $��3,0,3,0q�6%   66   �� 5#      页6   5$      5%      �� 5&   
   -   5        $��checkBox��5&      $��70��6#   $��27��6.   $��Center=�6:   $��启用设置��6X   $��202,50,20,73q�6g   6   �� 5'   
   -   5        $��70��5'      $��27��6   $��确定��6)   $��202,77,20,46q�66   5(      5)      5*      