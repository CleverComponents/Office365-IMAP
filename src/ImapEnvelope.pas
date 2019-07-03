unit ImapEnvelope;

interface

uses
  System.Classes, System.SysUtils;

type
  TImapEnvelope = class
  private
    FDate: TDateTime;
    FCcList: string;
    FFrom: string;
    FSubject: string;
    FBccList: string;
    FReplyTo: string;
    FMessageId: string;
    FInReplyTo: string;
    FSender: string;
    FToList: string;
    function ParseString(const ASource: string; var Index: Integer): string;
    function ParseEmailList(const ASource: string; var Index: Integer): string;
    function ParseEmail(const ASource: string; var Index: Integer): string;
  public
    procedure Parse(const ASource: string; var Index: Integer);
    procedure Clear;

    property Date: TDateTime read FDate write FDate;
    property Subject: string read FSubject write FSubject;
    property From: string read FFrom write FFrom;
    property Sender: string read FSender write FSender;
    property ReplyTo: string read FReplyTo write FReplyTo;
    property ToList: string read FToList write FToList;
    property CcList: string read FCcList write FCcList;
    property BccList: string read FBccList write FBccList;
    property InReplyTo: string read FInReplyTo write FInReplyTo;
    property MessageId: string read FMessageId write FMessageId;
  end;

implementation

uses
  clUtils, clMailHeader, clEncoder, clEmailAddress;

const
  NilLexem = 'NIL';
  ListSeparator: array[Boolean] of string = ('', ', ');

{ TImapEnvelope }

procedure TImapEnvelope.Clear;
begin
  Date := 0;
  FCcList := '';
  FFrom := '';
  FSubject := '';
  FBccList := '';
  FReplyTo := '';
  FMessageId := '';
  FInReplyTo := '';
  FSender := '';
  FToList := '';
end;

procedure TImapEnvelope.Parse(const ASource: string; var Index: Integer);
const
   EnvelopeLexem = 'ENVELOPE';
var
  ind: Integer;
begin
  Clear();

  ind := system.Pos(EnvelopeLexem, UpperCase(ASource), Index);
  if (ind < 1) then Exit;

  ind := system.Pos('(', ASource, ind + Length(EnvelopeLexem));
  if (ind < 1) then Exit;

  Inc(ind);

  Date := MimeTimeToDateTime(ParseString(ASource, ind));
  Subject := TclMailHeaderFieldList.DecodeField(ParseString(ASource, ind), '');
  From := ParseEmailList(ASource, ind);
  Sender := ParseEmailList(ASource, ind);
  ReplyTo := ParseEmailList(ASource, ind);
  ToList := ParseEmailList(ASource, ind);
  CcList := ParseEmailList(ASource, ind);
  BccList := ParseEmailList(ASource, ind);
  InReplyTo := ParseString(ASource, ind);
  MessageId := ParseString(ASource, ind);

  ind := system.Pos(')', ASource, ind);
  if (ind > 0) then
  begin
    Inc(ind);
    Index := ind;
  end;
end;

function TImapEnvelope.ParseString(const ASource: string; var Index: Integer): string;
var
  len, literalLen: Integer;
  next: PChar;
  isQuoted, isSafeChar, isLiteral: Boolean;
  nilInd: Integer;
begin
  Result := '';
  len := Length(ASource) - Index;
  if (len < 1) then Exit;

  isQuoted := False;
  isSafeChar := False;
  isLiteral := False;
  nilInd := 1;
  literalLen := -1;

  next := @ASource[Index];
  while (len > 0) do
  begin
    Inc(Index);

    case (next^) of
      '{':
        begin
          if (isQuoted) then
          begin
            Result := Result + next^;
          end else
          begin
            isLiteral := True;
            Result := '';
          end;
        end;
      '}':
        begin
          if (isQuoted) then
          begin
            Result := Result + next^;
          end else
          if (isLiteral) then
          begin
            literalLen := StrToIntDef(Trim(Result), 0);
            Result := '';
            if (len > 2) then
            begin
              Inc(Index, 2);
              Inc(next, 2);
              Dec(len, 2);
            end;
          end else
          begin
            Break;
          end;
        end;
      '"':
        begin
          if (isSafeChar) then
          begin
            Result := Result + next^;
            isSafeChar := False;
          end else
          if (isQuoted) then
          begin
            Break;
          end else
          if (isLiteral) then
          begin
            Result := Result + next^;
            if (literalLen > -1) then
            begin
              Dec(literalLen);
              if (literalLen = 0) then
              begin
                Break;
              end;
            end;
          end else
          begin
            isQuoted := True;
          end;
        end;
      '\':
        begin
          if (isSafeChar) then
          begin
            Result := Result + next^;
          end else
          if (isLiteral) then
          begin
            Result := Result + next^;
            if (literalLen > -1) then
            begin
              Dec(literalLen);
              if (literalLen = 0) then
              begin
                Break;
              end;
            end;
          end else
          begin
            isSafeChar := True;
          end;
        end;
      else
        begin
          if (isQuoted) then
          begin
            Result := Result + next^;
          end else
          if (isLiteral) then
          begin
            Result := Result + next^;
            if (literalLen > -1) then
            begin
              Dec(literalLen);
              if (literalLen = 0) then
              begin
                Break;
              end;
            end;
          end else
          begin
            if (NilLexem[nilInd] = UpperCase(next^)) then
            begin
              Inc(nilInd);
              if (nilInd > Length(NilLexem)) then
              begin
                Result := '';
                Break;
              end;
            end else
            begin
              nilInd := 1;
            end;
          end;
        end;
    end;

    Inc(next);
    Dec(len);
  end;
end;

function TImapEnvelope.ParseEmail(const ASource: string; var Index: Integer): string;
var
  addr: TclEmailAddressItem;
begin
  addr := TclEmailAddressItem.Create();
  try
    addr.Name := ParseString(ASource, Index);
    ParseString(ASource, Index);//addr-adl
    addr.Email := ParseString(ASource, Index) + '@' + ParseString(ASource, Index);

    Result := TclMailHeaderFieldList.DecodeEmail(addr.FullAddress, '');
  finally
    addr.Free();
  end;
end;

function TImapEnvelope.ParseEmailList(const ASource: string; var Index: Integer): string;
var
  len: Integer;
  next: PChar;
  nilInd, parCount: Integer;
begin
  Result := '';
  len := Length(ASource) - Index;
  if (len < 1) then Exit;

  parCount := 0;
  nilInd := 1;

  next := @ASource[Index];
  while (len > 0) do
  begin
    Inc(Index);

    case (next^) of
      '(':
        begin
          Inc(parCount);
          if (parCount = 2) then
          begin
            Result := Result + ListSeparator[Length(Result) > 0] + ParseEmail(ASource, Index);
            next := @ASource[Index];
            len := Length(ASource) - Index;
            Continue;
          end else
          if (parCount > 2) then
          begin
            Break;
          end;
        end;
      ')':
        begin
          Dec(parCount);
          if (parCount < 1) then
          begin
            Break;
          end;
        end
      else
        begin
          if (parCount = 0) and (NilLexem[nilInd] = UpperCase(next^)) then
          begin
            Inc(nilInd);
            if (nilInd > Length(NilLexem)) then
            begin
              Result := '';
              Break;
            end;
          end else
          begin
            nilInd := 1;
          end;
        end;
    end;

    Inc(next);
    Dec(len);
  end;
end;

end.
