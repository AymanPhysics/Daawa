   M S B A M L   `   `   `  ????@  <Daawa, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullQ-clr-namespace:WpfApplication1.EmployeeTrackerWpfApplication1.EmployeeTracker  5   1        WpfApplication1.MainWindow   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      <my-clr-namespace:WpfApplication1.EmployeeTracker     5   1     ? Title$  
MainWindow??6   $	??294??6   $	??618??6%    ? WindowState$ 	Maximized=?6j    ? WindowStartupLocation$ CenterScreen=?6?    ? WindowStyle$
 None=?6?   $??RightToLeft=?6?   $??Times New Roman)?6?   $
??Bold"?6?   $??16(?6   ? Triggers 5      .??5	      ? .+?5
   
   ? -   5        $??Grid1??5
      $
??0,39q?6    ? ColumnDefinitions 5      ?? 5      $
??Auto ?6#   ?? 5      $
??732* ?6#   .+?5      ?? -   5        $??Border1??5      $??1??6    ?? CornerRadius$ 3,3,3,3??6%   ???	#FFA5A5A5$??1,1,1,1q?6T   $??8,5,5,6q?6n   ._?5      ? .+?5      ?? -   5        $??TabControl1??5   !   $??	-1,0,0,-1q?64   $??0q?6G   6U   5      5      5      ,   &WpfApplication1.EmployeeTracker.Banner  5   
   -   5          Name$ Banner1??5   1   $??36??6   $??8,5,6,0q?6    $	??Top=?6B   1   +WpfApplication1.EmployeeTracker.BlankBanner  5   
   -   5          Name$ Banner2??5   6   $??35??6   $??8,0,6,6q?6%   $??Bottom=?6G   ?? 5   
   -   5        $??Button1??5   k   $$??الخروج من النظام??6   $??Right=?6,   	 ?? IsCancel	 .  $??0,0,12,6q?6Y   $	??136??6z   $??Bottom=?6?   $??34??6?      GlossyCloseButton#??[  6?   ??5      ?? 5      
 ?? EndPoint$
 0.5,1&?6&    ?? 
StartPoint$ 0.5,0&?67   .??5      ?  ? Color$	 Red??6#    ? Offset$ 0Q?6/   ? 5      $ 	#52000000??6#   $ 1Q?65   5      ??5      ?? 5      $
 0.5,1&?6&   $ 0.5,0&?67   .??5       ? $ White??6#   $ 0Q?61   ? 5!      $ 	#FF868686??6#   $ 1Q?65   5"      5$      ?? 5%   
   -   5        $??Button2??5%   h   $!??تغيير المستخدم??6   $
??Left=?6*   	 .  $??12,0,0,6q?6V   $??Bottom=?6w   $	??137??6?   #??[  6?   $??34??6?   ??5&      ?? 5'      $
 0.5,1&?6&   $ 0.5,0&?67   .??5(      ? $	 Red??6#   $ 0Q?6/   ? 5)      $ 	#52000000??6#   $ 1Q?65   5*      ??5,      ?? 5-      $
 0.5,1&?6&   $ 0.5,0&?67   .??5.      ? $ White??6#   $ 0Q?61   ? 5/      $ 	#FF868686??6#   $ 1Q?65   50      52      53      54      