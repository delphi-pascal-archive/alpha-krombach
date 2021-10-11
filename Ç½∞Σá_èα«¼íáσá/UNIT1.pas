UNIT Unit1;

INTERFACE

Uses
     Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
     Dialogs, Grids, StdCtrls, Math, Buttons, Menus, ComCtrls, ExtCtrls,
     ImgList;
     
TYPE
     TForm1= Class( Tform)
        Itm_About: TMenuItem;
        Itm_Calculer1: TMenuItem;
        Itm_Charger1: TMenuItem;
        Itm_Editer1: TMenuItem;
        Itm_Fichier1: TMenuItem;
        Itm_Jouer1: TMenuItem;
        Itm_L_Alpha: TMenuItem;
        Itm_Le_programme: TMenuItem;
        Itm_Les_calculs: TMenuItem;
        Itm_Nouveau1: TMenuItem;
        Itm_Parametrer1: TMenuItem;
        Itm_Quitter1: TMenuItem;
        Itm_Remplir1: TMenuItem;
        Itm_Remplir2: TMenuItem;
        Itm_Remplir3: TMenuItem;
        Itm_Remplir4: TMenuItem;
        Itm_Remplir5: TMenuItem;
        Itm_Sauver: TMenuItem;
        Itm_SauverSous: TMenuItem;
        Itm_sep2: TMenuItem;
        Itm_sep3: TMenuItem;
        Itm_sep4: TMenuItem;
        Itm_sep5: TMenuItem;
        Label1: TLabel;
        Label2: TLabel;
        Lbl_Alpha: TLabel;
        MainMenu1: TMainmenu;
        OpenDialog1: TOpenDialog;
        Richedit1: TRichedit;
        SaveDialog1: TSaveDialog;
        Sep1: TMenuItem;
        Stringgrid1: TStringgrid;
        Btn_OK: TButton;
        
        
        PROCEDURE Cacher_les_labels;
        PROCEDURE Calculer(Sender: Tobject);
        PROCEDURE Colorier(Sender: Tobject; C,R: Longint;Rect: Trect; State: Tgriddrawstate);
        PROCEDURE Demander_les_parametres(Sender:Tobject);
        PROCEDURE Editer(Sender:Tobject);
        PROCEDURE Expliquer_l_Alpha(Sender:Tobject);
        PROCEDURE FormCreate(Sender:Tobject);
        PROCEDURE FormKeyDown(Sender: Tobject; Var Key: Word;Shift: Tshiftstate);
        PROCEDURE FormKeyPress(Sender: Tobject; Var Key: Char);
        PROCEDURE FormClose(Sender: Tobject);
        PROCEDURE Itm_Nouveau1Click(Sender: Tobject);
        PROCEDURE Mettre_les_composants_en_bonne_place;
        PROCEDURE Montrer_les_labels;
        PROCEDURE Ouvrir(Sender:Tobject);
        PROCEDURE Passer_en_ecran_de_saisie;
        PROCEDURE Placer_le_bouton_OK;
        PROCEDURE Preparer_remplissage;
        PROCEDURE Presenter_calculs(Sender:Tobject);
        PROCEDURE Presenter_le_programme(Sender:Tobject);
        PROCEDURE Remplir_2(Sender:Tobject);
        PROCEDURE Remplir_3(Sender:Tobject);
        PROCEDURE Remplir_4(Sender:Tobject);
        PROCEDURE Remplir_5(Sender:Tobject);
        PROCEDURE Remplir_de_valeurs_aleatoires(Sender:Tobject);
        PROCEDURE Retablir_les_titres;
        PROCEDURE Sauver(Sender:Tobject);
        PROCEDURE Sauver_Sous(Sender:Tobject);
        PROCEDURE Stringgrid1Drawcell(Sender: Tobject; Col, Row: Longint; Rect: Trect; State: Tgriddrawstate);
        PROCEDURE Btn_OKClick(Sender: Tobject);
    procedure StringGrid1Click(Sender: TObject);
        
      Private
        { Déclarations privées }
      Public
        { Déclarations publiques }
     End;
     
VAR
    Alpha: Extended;
    Dir_initial:String;
    F1: Text;
    Fichier_en_cours: String;
    Form1: TForm1;
    K: BYTE;
    Nb_cols_maxi,Nb_rows_maxi: BYTE;
    Nb_personnes: Word;
    Premiere_fois: BOOLEAN;
    
    
IMPLEMENTATION

{$R *.dfm}

FUNCTION Ltrim(S:String):String;
Begin
   WHILE (Length(S)>1) AND (S[1]=' ') DO S:=Copy(S,2,255);
   Ltrim:=S;
End;

FUNCTION Rtrim(S:String):String;
Begin
   WHILE (Length(S)>0) AND (S[Length(S)]=' ') DO SetLength(S,Length(S)-1);
   Rtrim:=S;
End;

FUNCTION Alltrim(S:String):String;
Begin
   Alltrim:=Ltrim(Rtrim(S));
End;

PROCEDURE TForm1.Montrer_les_labels;
Begin
   Label1.Visible:=TRUE;
   Lbl_Alpha.Visible:=TRUE;
End;

PROCEDURE TForm1.Cacher_les_labels;
Begin
   Label1.Visible:=FALSE;
   Lbl_Alpha.Visible:=FALSE;
End;


PROCEDURE TForm1.Retablir_les_titres;
VAR I,P: BYTE;
Begin
   WITH Stringgrid1 DO
   Begin
      Options:=Options+[GoEditing];
      FOR I:=1 TO ColCount-2 DO
      Begin
         Cells[I,0]:='It. '+InttoStr(I);
         ColWidths[I]:=DefaultColWidth;
      End;
      
      FOR P:=1 TO RowCount-1 DO
      Begin
         Cells[0,P]:='Pers nº'+InttoStr(P);
         RowHeights[P]:=Defaultrowheight;
      End;
   End;
End;

PROCEDURE TForm1.FormClose(Sender: Tobject);
Begin
   Application.Terminate;
   HALT;
End;

PROCEDURE TForm1.Placer_le_bouton_OK;
Begin
   WITH Btn_OK DO
   Begin
      Visible:=TRUE;
      Top:=Richedit1.Top+Richedit1.Height-Height;
      Left:=Richedit1.Left + (Richedit1.Width-Width)Div 2;
   End;
End;


PROCEDURE TForm1.Presenter_calculs(Sender:Tobject);
Begin
   WITH Richedit1 DO
   Begin
      Top:= 1;
      Left:= 600;
      Width:= 370;
      Height:= 670;
      Clear;
      Lines.Add('                        Les calculs ne sont pas très durs !');
      Lines.Add('                                                   ----     ');
      Lines.Add(' On part de la formule:');
      Lines.Add('');
      Lines.Add('(k/(k-1))(1 - (somme des variances / variance des sommes))');
      Lines.Add('');
      Lines.Add('   k représente le nombre d''items   (ou nombre de questions posées)');
      Lines.Add('   Si k est assez élevé, k/(k-1) est proche de 1 et donc');
      Lines.Add(' ce premier terme n''a pas beaucoup d''influence.');
      (*
      Lines.Add('   Si k est égal à 10, ce terme vaut 0,9, c''est presque 1');
      Lines.Add('   Par contre, si k=2, ce terme vaut 2.');
      *)
      Lines.Add('   Si k=1, division par zéro, impossible de faire le calcul!');
      Lines.Add('');
      Lines.Add('   Les sommes par personne sont sur la derniŠre colonne du tableau.');
      Lines.Add('   Elles sont en bleu, leur variance est en vert.');
      Lines.Add('');
      Lines.Add('   Les variances des items sont sur la dernière ligne du tableau.');
      Lines.Add('   Elles sont en rose, leur somme est en rouge.');
      Lines.Add('');
      Lines.Add('  On divise la somme des variances par la variance des sommes.'+
      '  c''est-à-dire la case rouge par la case verte.');
      Lines.Add('');
      Lines.Add('    On retranche le résultat de 1 et enfin, on multiplie ce');
      Lines.Add('   dernier résultat par le premier terme (k/(k-1))');
      Lines.Add('    On obtient souvent un résultat compris entre 0,5 et 1.');
      Lines.Add('   Mais ce n''est pas toujours le cas. On atteint parfois');
      Lines.Add('   de grandes valeurs, positives ou négatives.');
      Lines.Add('');
      Lines.Add('  Si la somme des réponses est constante, exemple:');
      Lines.Add('  Personne 1:        1   2   3     Somme=6');
      Lines.Add('  Personne 2:        3   2   1     Somme=6');
      Lines.Add('  Personne 3:        0   3   3     Somme=6');
      Lines.Add('  Alors, la variance des sommes vaut zéro et on ne peut');
      Lines.Add(' pas calculer l''alpha de Cronbach. Mais si on modifie');
      Lines.Add(' légèrement une des valeurs, (on remplace un 3 par 2.99) cela devient possible.');
      Lines.Add(' Mais attention, cette petite modification de données modifie considérablement le résultat.');
      Lines.Add(' Voyez par exemple la différence si on remplace 3 par 2.99 puis par 2.98 !');
      Visible:=TRUE;
   End;
   Placer_le_bouton_OK;
End;

PROCEDURE TForm1.Expliquer_l_Alpha(Sender:Tobject);
VAR
    Fichier_source: String;
    Ligne:String;
    Longueur_ligne_maxi:BYTE;
    Nb_lignes: Word;
    
Begin
   Fichier_source:=Dir_initial +'\Cronbach.txt';
   AssignFile(F1,Fichier_source);
   {$I-} Reset(F1); {$I+}
   IF IOResult=0 THEN
   Begin
      WITH Richedit1 DO
      Begin
         Longueur_ligne_maxi:=51;
         Nb_lignes:=0;
         WHILE NOT EOF(F1) DO
         Begin
            Inc(Nb_lignes);
            Readln(F1,Ligne);
            IF (Length(Ligne)-2)> Longueur_ligne_maxi THEN
               Longueur_ligne_maxi:=Length(Ligne)-2;
         End;
         CloseFile(F1);
         
         Clear;
         Lines.LoadFromFile(Fichier_source);
         Lines.Insert(0,' --- Ceci est le texte du fichier Cronbach.txt --- ');
         
         Height:=Nb_lignes *12  +50 ;
         IF Height>600 THEN Height:=600;
         Width:=Longueur_ligne_maxi *6;
         IF Width>1000 THEN Width:=1000;
         
         Top:=10;
         Left:=Form1.Width-20-Width;
         Visible:=TRUE;
      End;
   End
   ELSE
   Begin
      WITH Richedit1 DO
      Begin
         Top:=10;
         Left:=648;
         Width:= 310;
         Height:=180;
         Clear;
         Lines.Add('                                  Alpha de Cronbach');
         Lines.Add('                                                 ----     ');
         Lines.Add('   Si, dans le même répertoire que ce programme,');
         Lines.Add(' il existait un fichier nommé Cronbach.txt, c''est le');
         Lines.Add(' texte de ce fichier que vous auriez sous les yeux');
         Lines.Add(' et il vous expliquerait ce qu''exprime l''alpha de');
         Lines.Add(' Cronbach.');
         Lines.Add('');
         Lines.Add('  Ça viendra peut-être un jour !');
         Lines.Add('');
         Visible:=TRUE;
      End;
   End;
   Placer_le_bouton_OK;
End;

PROCEDURE TForm1.Presenter_le_programme(Sender:Tobject);
Begin
   WITH Richedit1 DO
   Begin
      Top:=10;
      Left:=648;
      Width:= 160;
      Height:=170;
      Clear;
      Lines.Add('        Cronbach''s alpha');
      Lines.Add('                       ----     ');
      Lines.Add('             by Fredelem ');
      Lines.Add('           fredelem@sfr.fr ');
      Lines.Add('');
      Lines.Add('               July 2010   ');
      Lines.Add('');
      Lines.Add('         Distribute freely');
      Lines.Add('');
      Visible:=TRUE;
   End;
   Placer_le_bouton_OK;
End;

PROCEDURE TForm1.Calculer(Sender:Tobject);
Label
      Fin;
VAR
    Alpha: Extended;
    Erreur: INTEGER;
    I,P: BYTE;
    Nb_chiffres_apres_virgule: BYTE;
    Valeur,Somme,Moyenne:Extended;
    Somme_des_variances: Extended;
    Variance_des_sommes: Extended;
    S:String[10];

Begin
   Richedit1.Visible:=FALSE;
   Label2.Visible:=FALSE;
   IF (Stringgrid1.Colcount<3) THEN
   Begin
      Showmessage('Désolé, il faut au moins 2 items.');
      EXIT;
   End;
   
   WITH Stringgrid1 DO
   Begin
      {On adapte la stringgrid aux nouveaux paramètres}
      ColCount:=K +3;
      Cells[ColCount-2,0]:='';
      Cells[ColCount-1,0]:='Somme';
      
      RowCount:=Nb_personnes +5;
      RowHeights[Nb_personnes +1]:=10;
      ColWidths[ColCount-2]:=10;
      ColWidths[ColCount-1]:=50;
      Cells[0,Nb_personnes+1]:='';
      Cells[0,RowCount-3]:='Somme';
      Cells[0,RowCount-2]:='Moyenne';
      Cells[0,RowCount-1]:='Variance';
      
      {On lit et on calcule}
      FOR P:=1 TO Nb_personnes DO
      Begin
         Somme:=0;
         FOR I :=1 TO K DO
         Begin
            Cells[I ,P]:=Alltrim(Cells[I ,P]);
            Val(Cells[I ,P],Valeur,Erreur);
            IF Erreur>0 THEN
            Begin
               Showmessage('La cellule de l''item nª'+Inttostr(I )+' personne nº'+Inttostr(P)+
               ' ne contient pas une valeur numérique correcte.');
               Col:=I ;
               Row:=P;
               Options:=Options +[GoEditing];
               SetFocus;
               EXIT;
            End;
            Somme:=Somme +Valeur;
         End;
         Str(Somme:1:2,S);
         // S représente la somme des réponses pour la personne P. On l'affiche
         Cells[K +2,P]:=S;
      End;
      
      {Item par item, on calcule les 3 lignes: somme, moyenne, variance}
      FOR I:=1 TO K +2 DO
      Begin
         IF I<>(K +1) THEN
         Begin
            // On commence par calculer la somme de l'item et par l'afficher
            Somme:=0;
            FOR P:=1 TO Nb_personnes DO
               Somme:=Somme +StrtoFloat(Cells[I ,P]);
            Str(Somme:1:2,S);
            Cells[I ,Nb_personnes+2]:=S;
            
            // Puis on calcule la moyenne de l'item et on l'affiche
            Moyenne:= Somme /Nb_personnes;
            Str(Moyenne:1:2,S);
            Cells[I ,Nb_personnes +3]:=S;
            
            // Puis on calcule la variance des items et des sommes et on l'affiche
            Somme:=0;
            FOR P:=1 TO Nb_personnes DO
            Begin
               Somme:=Somme +Power((Moyenne-Strtofloat(Cells[I ,P])),2);
               IF ColWidths[I ]<(Length(Cells[I ,P])*(Font.Size)) THEN
                  ColWidths[I ]:=(Length(Cells[I ,P])*(Font.Size));
            End;
            Variance_des_sommes:= Somme /Nb_personnes;
            Str(Variance_des_sommes:1:2,S);
            Cells[I ,Nb_personnes +4]:=S;
            
            // Si la cellule est trop étroite, on l'agrandit un peu
            P:=Nb_personnes +3;
            IF ColWidths[I ]<(Length(Cells[I ,P])*(Font.Size)) THEN
               ColWidths[I ]:=(Length(Cells[I ,P])*(Font.Size));
         End;
      End;
      
      {On calcule alpha}
      Somme_des_variances:=0;
      FOR I :=1 TO K DO
         Somme_des_variances:=Somme_des_variances +StrtoFloat(Cells[I ,Nb_personnes +4]);
      
      Cells[K+2,Nb_personnes +2]:=FloattoStr(Variance_des_sommes);
      Cells[K+2,Nb_personnes +3]:='';
      
      IF Variance_des_sommes=0 THEN
      Begin
         Showmessage('Impossible de calculer l''alpha de Cronbach'+#13+
         'car la variance de la somme (case verte) vaut 0'+#13+
         'Il faudrait diviser par zéro');
         GOTO Fin;
      End
      ELSE
         Cells[K +2,Nb_personnes +4]:=FloattoStr(Somme_des_variances);
      
      
      Alpha:=(K /(K-1)) * (1-(Somme_des_variances/Variance_des_sommes));
      
      Montrer_les_labels;
      WITH Lbl_Alpha DO
      Begin
         Nb_chiffres_apres_virgule:=12;
         BringtoFront;
         Autosize:=TRUE;
         Repeat
            Str(Alpha:1:Nb_chiffres_apres_virgule,S);
            Valeur:=StrtoFloat(S);
            IF Valeur=Trunc(Valeur) THEN
               Inc(Nb_chiffres_apres_virgule)
            ELSE
            Begin
               Str(Alpha:1:Nb_chiffres_apres_virgule,S);
               Break;
            End;
         Until(Nb_chiffres_apres_virgule>=12);
         Caption:= S;
         Autosize:=FALSE;
         Height:=Label1.Height;
      End;
   End;
   Fin:
   WITH Stringgrid1 DO
      Options:=Options-[GoEditing];
   // On simule un click de souris (pour éviter que la 1ère cellule ne soit modifiable)
   SetCursorPos(5,5);
   SendMessage(Stringgrid1.Handle,Wm_lbuttondown,0,0);
   // On rend à nouveau visibles les options du menu
   Itm_Editer1.Enabled:=TRUE;
   Label2.Visible:=TRUE;
End;

PROCEDURE TForm1.Colorier(Sender: Tobject; C,R : Longint; Rect: Trect; State: Tgriddrawstate);
Begin
   {Les séparations doivent apparaître sur fond gris}
   IF (C=K +1) OR (R =Nb_personnes +1) THEN
   Begin
      Stringgrid1.Canvas.Brush.Color:=Clgray;;
      Stringgrid1.Canvas.Fillrect(Rect);
   End
   
   ELSE {Les cellules après la colonne de séparation seront en Lime}
   IF (C=K +2) THEN
   Begin
      IF (R =Nb_personnes +2) THEN
      Begin
         Stringgrid1.Canvas.Brush.Color:=ClLime;
         Stringgrid1.Canvas.Fillrect(Rect);
      End
      ELSE
      IF (R =Nb_personnes +4) THEN
      Begin
         Stringgrid1.Canvas.Brush.Color:=$002A2fff;
         Stringgrid1.Canvas.Fillrect(Rect);
      End
      ELSE
      Begin
         Stringgrid1.Canvas.Brush.Color:=Claqua;
         Stringgrid1.Canvas.Fillrect(Rect);
      End;
   End

   (*
   ELSE {Les cellules "Sommes" seront en jaune}
   IF (R =Nb_personnes +2) THEN
   Begin
      Stringgrid1.Canvas.Brush.Color:=ClYellow;
      Stringgrid1.Canvas.Fillrect(Rect);
   End
   *)

   ELSE {Les cellules "Variances"seront en tose}
   IF (R =Nb_personnes +4) THEN
   Begin
      Stringgrid1.Canvas.Brush.Color:=$00D4d4fe;
      Stringgrid1.Canvas.Fillrect(Rect);
   End
   
   ELSE {Toutes les autres cellules doivent apparaître sur fond blanc}
   Begin
      Stringgrid1.Canvas.Brush.Color:=ClWhite;;
      Stringgrid1.Canvas.Fillrect(Rect);
   End;
   
   
   // On fixe la couleur des lignes et colonnes fixes
   IF (Gdfixed In State) THEN
   Begin
      Stringgrid1.Canvas.Brush.Color:=$00C8b9c7;
      Stringgrid1.Canvas.Fillrect(Rect)
   End;
   
   // On fixe la couleur de la fonte
   Stringgrid1.Canvas.Font.Color:=ClBlack;
   Stringgrid1.Canvas.Textout(Rect.Left +2,Rect.Top +2,Stringgrid1.Cells[C,R ]);
End;

PROCEDURE TForm1.Itm_Nouveau1Click(Sender:Tobject);
Begin
   Cacher_les_labels;
   WITH Lbl_Alpha DO
   Begin
      Caption:='';
      Autosize:=TRUE;
   End;
   Passer_en_ecran_de_saisie;
   K:=0;
   Nb_personnes:=0;
   Demander_les_parametres(Nil);
   Retablir_les_titres;
   IF Premiere_fois AND (Stringgrid1.Colcount>2) AND (K>1) THEN
   Begin
      Showmessage('Saisissez les valeurs dans le tableau puis cliquez sur "Calculer"');
      Premiere_fois:=FALSE;
   End;
   WITH Stringgrid1 DO
   Begin
      Col:=1;
      Row:=1;
      SetFocus;
   End;
   Itm_Editer1.Enabled:=TRUE;
   Itm_Parametrer1.Enabled:=TRUE;
   Itm_Calculer1.Enabled:=TRUE;
   Itm_Jouer1.Enabled:=TRUE;
   Itm_Sauver.Enabled:=TRUE;
   Itm_SauverSous.Enabled:=TRUE;
End;

PROCEDURE TForm1.FormCreate(Sender:Tobject);
VAR I,P:BYTE;
Begin
   Dir_initial:=ExtractFilePath(Paramstr(0));
   SaveDialog1.InitialDir:=Dir_initial;
   OpenDialog1.InitialDir:=Dir_initial;
   DecimalSeparator:='.';
   Premiere_fois:=TRUE;
   WITH Label2 DO
   Begin
      Top:=650;
      Left:=10;
      Visible:=FALSE;
   End;
   WITH Form1 DO
   Begin
      Top:=0;
      Left:=0;
      Width:=Screen.Width;
      Height:=Screen.Height;
   End;
   
   WITH Stringgrid1 DO
   Begin
      Nb_rows_maxi:=RowCount;
      Nb_cols_maxi:=ColCount;
      FOR I:=1 TO ColCount-2 DO Cells[I,0]:='It. '+InttoStr(I);
      Cells[ColCount,0]:='Somme';
      FOR P:=1 TO RowCount-3 DO Cells[0,P]:='Pers nº'+InttoStr(P);
      Cells[0,RowCount-1]:='Variance';
      ColWidths[0]:=60;
      Color:=ClTeal;
      ColCount:=0;
      RowCount:=0;
   End;
   
   K:=0;
   Nb_personnes:=0;
   Mettre_les_composants_en_bonne_place;
End;

PROCEDURE TForm1.Mettre_les_composants_en_bonne_place;
{Donne aux composants la position et la taille qui conviennent}
Begin
   WITH Form1 DO
   Begin
      Left:=0;
      Top:=0;
      Width:=1024;
      Height:=748;
   End;
   WITH Label1 DO
   Begin
      Left:=360;
      Top:=642;
      Width:=323;
      Height:=33;
   End;
   WITH Lbl_Alpha DO
   Begin
      Left:=682;
      Top:=642;
      Width:=7;
      Height:=24;
   End;
   WITH Stringgrid1 DO
   Begin
      Left:=0;
      Top:=0;
      Width:=337;
      Height:=105;
   End;
End;

PROCEDURE TForm1.Passer_en_ecran_de_saisie;
VAR I,P: BYTE;
Begin
   WITH Stringgrid1 DO
   Begin
      Options:=Options+[GoEditing];
      Visible:=TRUE;
      Color:=ClSkyBlue;
      Width:=Form1.Width-10;
      Height:=640;
      ColCount:=2;
      RowCount:=2;
      ColWidths[0]:=70;
      FixedRows:=1;
      FixedCols:=1;
      FixedColor:=ClYellow;
      FOR I:=1 TO Nb_cols_maxi DO
      FOR P:=1 TO Nb_rows_maxi-1 DO
         Cells[I,P]:='';
      Col:=1;
      Row:=1;
      SetFocus;
   End;
End;

PROCEDURE TForm1.Editer(Sender:Tobject);
Begin
   Itm_Editer1.Enabled:=FALSE;
   Richedit1.Visible:=FALSE;
   Cacher_les_labels;
   WITH Stringgrid1 DO
   Begin
      Options:=Options+[GoEditing];
      ColCount:=K+1;
      RowCount:=Nb_personnes+1;
      IF (K<2) OR (Nb_personnes<1) THEN
      Begin
         Passer_en_ecran_de_saisie;
         Demander_les_parametres(Nil);
      End;
      Retablir_les_titres;
   End;
End;

PROCEDURE TForm1.Demander_les_parametres(Sender:Tobject);
VAR
    Erreur: INTEGER;
    I ,P: BYTE;
    Reponse:String;
Begin
   Richedit1.Visible:=FALSE;
   Label2.Visible:=FALSE;
   Cacher_les_labels;
   Itm_Editer1.Enabled:=FALSE;
   
   WITH Stringgrid1 DO {On efface les restes du calcul précédent}
   Begin
      Options:=Options +[GoEditing];
      IF RowCount>Nb_personnes +1 THEN
      FOR I :=1 TO ColCount DO
      Begin
         Cells[I ,Nb_personnes +1]:='';
         Cells[I ,Nb_personnes +2]:='';
         Cells[I ,Nb_personnes +3]:='';
      End;
      IF ColCount>K +1 THEN
      FOR P:=1 TO RowCount DO
      Begin
         Cells[K +1,P]:='';
         Cells[K +2,P]:='';
         Cells[K +3,P]:='';
      End;
   End;
   
   Reponse:='5'; IF Nb_personnes>0 THEN Str(Nb_personnes,Reponse);
   Repeat
      IF InputQuery('','Combien de personnes ont été interrogées ?',Reponse) THEN
      Begin
         Val(Reponse,Nb_personnes,Erreur);
         IF Erreur>0 THEN Showmessage(Reponse +' n''est pas un nombre entier correct')
         ELSE
         IF (Nb_personnes<1) THEN Showmessage('Le nombre de personnes doit être au moins égal à 1');
      End
      ELSE
         EXIT;
   Until Error=0;

   Reponse:='5'; IF K>0 THEN Str(K,Reponse);
   Repeat
      IF InputQuery('','Nombre d''items ? (= Nb de questions posées)',Reponse) THEN
      Begin
         Val(Reponse,K,Erreur);
         IF Erreur>0 THEN Showmessage(Reponse +' n''est pas un nombre entier correct')
         ELSE
         IF (K<2) THEN Showmessage('Le nombre d''items doit être au moins égal à 2');
      End
      ELSE
         EXIT;
   Until(K>=2);
   
   WITH Stringgrid1 DO
   Begin
      ColCount:=K +1;
      RowCount:=Nb_personnes +1;
      Retablir_les_titres;
   End;
End;

PROCEDURE TForm1.Preparer_remplissage;
VAR
    I ,P: BYTE;
Begin
   Cacher_les_labels;
   WITH Stringgrid1 DO
      Options:=Options +[GoEditing];
   Itm_Editer1.Enabled:=FALSE;

   IF (K<2) OR (Nb_personnes<1) THEN
   Begin
      Passer_en_ecran_de_saisie;
      Demander_les_parametres(Nil);
   End;

   IF (K<2) OR (Nb_personnes<1) THEN EXIT;

   WITH Stringgrid1 DO {On efface les restes du calcul précédent}
   Begin
      Options:=Options +[GoEditing];
      IF RowCount>Nb_personnes +1 THEN
      FOR I :=1 TO ColCount DO
      Begin
         Cells[I ,Nb_personnes +1]:='';
         Cells[I ,Nb_personnes +2]:='';
         Cells[I ,Nb_personnes +3]:='';
      End;
      IF ColCount>K +1 THEN
      FOR P:=1 TO RowCount DO
      Begin
         Cells[K +1,P]:='';
         Cells[K +2,P]:='';
         Cells[K +3,P]:='';
      End;
      ColCount:=K +1;
      RowCount:=Nb_personnes +1;
      Retablir_les_titres;
   End;
End;


PROCEDURE TForm1.Remplir_de_valeurs_aleatoires(Sender:Tobject);
VAR
    Erreur: INTEGER;
    I ,P: BYTE;
    Val_maxi: Extended;
    V: String;
Begin
   Preparer_remplissage;
   IF (K<2) OR (Nb_personnes<1) THEN EXIT;
   
   Val_maxi:=10;
   V:=FloattoStr(Val_maxi);
   Randomize;
   Repeat
      IF InputQuery('','Valeur maximum ?',V) THEN
      Begin
         Val(V,Val_maxi,Erreur);
         IF Erreur=0 THEN
         Begin
            FOR I :=1 TO K DO
            FOR P:=1 TO Nb_personnes DO
               Stringgrid1.Cells[I ,P]:=InttoStr(Random(Trunc(Val_maxi +1)));
         End;
      End
   Until(Error=0);
   Itm_Sauver.Enabled:=TRUE;
   Itm_SauverSous.Enabled:=TRUE;
End;


PROCEDURE TForm1.Remplir_2(Sender:Tobject);
VAR
    Chaine: String;
    I ,P: BYTE;
    Erreur: INTEGER;
    Valeur: Extended;
Begin
   Preparer_remplissage;
   IF (K<2) OR (Nb_personnes<1) THEN EXIT;
   Repeat
      IF InputQuery('','Remplir avec quel nombre ? ',Chaine) THEN
      Begin
         Val(Chaine,Valeur,Erreur);
         IF Erreur=0 THEN
         Begin
            FOR I:=1 TO K DO
            FOR P:=1 TO Nb_personnes DO
               IF Valeur<100 THEN
                  Stringgrid1.Cells[I ,P]:=FloattoStr(Valeur)
               ELSE
                  Stringgrid1.Cells[I ,P]:=Chaine;
         End
         ELSE
            Showmessage(Chaine+' n''est pas un nombre !');
      End
      ELSE
         EXIT;
   Until Erreur=0;
   Itm_Sauver.Enabled:=TRUE;
   Itm_SauverSous.Enabled:=TRUE;
End;

PROCEDURE TForm1.Remplir_3(Sender:Tobject);
VAR
    I ,P: BYTE;
Begin
   Preparer_remplissage;
   IF (K<2) OR (Nb_personnes<1) THEN EXIT;
   FOR I :=1 TO K DO
   FOR P:=1 TO Nb_personnes DO
      Stringgrid1.Cells[I ,P]:=InttoStr(P);;
   Itm_Sauver.Enabled:=TRUE;
   Itm_SauverSous.Enabled:=TRUE;
End;


PROCEDURE TForm1.Remplir_4(Sender:Tobject);
VAR
    I ,P: BYTE;
Begin
   Preparer_remplissage;
   IF (K<2) OR (Nb_personnes<1) THEN EXIT;
   FOR I :=1 TO K DO
   FOR P:=1 TO Nb_personnes DO
      Stringgrid1.Cells[I ,P]:=InttoStr(I );;
   Itm_Sauver.Enabled:=TRUE;
   Itm_SauverSous.Enabled:=TRUE;
End;

PROCEDURE TForm1.Remplir_5(Sender:Tobject);
VAR
    I ,P: BYTE;
Begin
   Preparer_remplissage;
   IF (K<2) OR (Nb_personnes<1) THEN EXIT;
   WITH Stringgrid1 DO
   FOR I :=1 TO K DO
   FOR P:=1 TO Nb_personnes DO
   Begin
      Cells[I ,P]:=FloattoStr(10*Strtofloat(Cells[I ,P]));
      IF ColWidths[I ]<(Length(Cells[I ,P])*(Font.Size)) THEN
         ColWidths[I ]:=(Length(Cells[I ,P])*(Font.Size));
   End;
   Itm_Sauver.Enabled:=TRUE;
   Itm_SauverSous.Enabled:=TRUE;
End;

PROCEDURE TForm1.Ouvrir(Sender:Tobject);
VAR
    V: Word;
    IOR:INTEGER;
    Ligne: String;
    
Begin
   Itm_Editer1.Enabled:=TRUE;
   Itm_Parametrer1.Enabled:=TRUE;
   Itm_Calculer1.Enabled:=TRUE;
   Itm_Jouer1.Enabled:=TRUE;
   Cacher_les_labels;
   Itm_Editer1.Enabled:=FALSE;
   Passer_en_ecran_de_saisie;
   
   Repeat
      WITH OpenDialog1 DO
      Begin
         Filename:='*.CRN';
         IF OpenDialog1.Execute THEN
         Begin
            Fichier_en_cours:=Filename;
            IF Fichier_en_cours='' THEN EXIT;
            AssignFile(F1,Fichier_en_cours);
            {$I-} Reset(F1); {$I+}
            IOR:=IOResult;
            IF IOR<>0 THEN
            Begin
               Showmessage('Fichier non trouvé.');
               EXIT;
            End;
            K:=0; Nb_personnes:=0;
            WHILE NOT EOF(F1) DO
            Begin
               Repeat
                  IF EOF(F1) THEN Break;
                  Readln(F1,Ligne);
               Until Alltrim(Ligne)<>'';
               
               IF Ligne<>'' THEN
               Begin
                  Inc(Nb_personnes);
                  K:=0;
                  WHILE Ligne<>'' DO
                  Begin
                     V:=Pos(',',Ligne); IF V=0 THEN V:=Length(Ligne)+1;
                     Inc(K);
                     Stringgrid1.Cells[K,Nb_personnes]:=Alltrim(Copy(Ligne,1,V-1));
                     Ligne:=Alltrim(Copy(Ligne,V+1,255));
                  End;
               End;
               WITH Stringgrid1 DO
               Begin
                  ColCount:=K+1;
                  RowCount:=Nb_personnes+1;
               End;
               
            End;
            CloseFile(F1);
            WITH Stringgrid1 DO
               Retablir_les_titres;
         End
         ELSE
            EXIT;
      End
   Until IOR=0;
   Itm_Sauver.Enabled:=TRUE;
   Itm_SauverSous.Enabled:=TRUE;
   
End;

PROCEDURE TForm1.Sauver_Sous(Sender:Tobject);
Begin
   WITH SaveDialog1 DO
   Begin
      Filename:='*.CRN';
      IF SaveDialog1.Execute THEN
      Begin
         Fichier_en_cours:=Filename;
         Sauver(Nil);
      End;
   End;
End;

PROCEDURE TForm1.Sauver(Sender:Tobject);
VAR
    I,P: BYTE;
Begin
   IF Fichier_en_cours='' THEN
      Sauver_Sous(Nil)
   ELSE
   Begin
      AssignFile(F1,Fichier_en_cours);
      REWRITE(F1);
      FOR P:=1 TO Nb_personnes DO
      Begin
         FOR I:=1 TO K DO
         Begin
            Write(F1,Stringgrid1.Cells[I,P]);
            IF I<K THEN Write(F1,',');
         End;
         WriteLn(F1);
      End;
      CloseFile(F1);
   End;
End;

PROCEDURE TForm1.FormKeyDown(Sender: Tobject; Var Key: Word;Shift: Tshiftstate);
Begin
   WITH Stringgrid1 DO
   Begin
      IF (Col=1) AND (Key=Vk_left) THEN Key:=0;
      
      IF (Key=Vk_right) AND (Col=Colcount-1) THEN
      Begin
         IF (Row<Rowcount-1)  THEN
         Begin
            Row:=Row+1;
            Col:=1;
            Key:=0;
         End
         ELSE
            Calculer(Nil);
      End;
      
      IF (Key=Vk_down) AND (Row=Rowcount-1) THEN
      Begin
         IF (Col<Colcount-1)  THEN
         Begin
            Col:=Col+1;
            Row:=1;
            Key:=0;
         End
         ELSE
            Calculer(Nil);
      End;
   End;
End;

PROCEDURE TForm1.FormKeyPress(Sender: Tobject; Var Key: Char);
Begin
   WITH Stringgrid1 DO
   Begin
      IF (Key=#13) AND (Col=Colcount-1) THEN
      Begin
         IF (Row<Rowcount-1) THEN
         Begin
            Row:=Row+1;
            Col:=1;
            Key:=#0;
         End
         ELSE
            Calculer(Sender);
      End
   End;
End;

PROCEDURE TForm1.Stringgrid1Drawcell(Sender: Tobject; Col, Row: Longint; Rect: Trect; State: Tgriddrawstate);
Begin
   Colorier(Sender,Col, Row,Rect,State);
End;

PROCEDURE TForm1.Btn_OKClick(Sender: Tobject);
Begin
   Richedit1.Visible:=FALSE;
   Btn_OK.Visible:=FALSE;
End;

PROCEDURE TForm1.StringGrid1Click(Sender: TObject);
Begin
   if not (GoEditing in Stringgrid1.options) then
      Showmessage('Pour pouvoir modifier une case,'+#13+'    cliquez d''abord sur "Editer"');
End;

End.
