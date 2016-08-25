package doclet.docx;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.List;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import com.sun.javadoc.ClassDoc;
import com.sun.javadoc.ExecutableMemberDoc;
import com.sun.javadoc.MemberDoc;
import com.sun.javadoc.MethodDoc;
import com.sun.javadoc.PackageDoc;
import com.sun.javadoc.ParamTag;
import com.sun.javadoc.Parameter;
import com.sun.javadoc.RootDoc;
import com.sun.javadoc.Tag;
import com.sun.javadoc.ThrowsTag;
import com.sun.javadoc.Type;

/**
 * Microsoft Word 形式の Javadoc ドキュメントを作成する処理を提供します。
 */
public class DocumentBuilder {

	/**
	 * Word 文書
	 */
	private XWPFDocument word;

	/**
	 * Javadoc のルートドキュメント
	 */
	private RootDoc root;

	/**
	 * 出力済のパッケージを記憶するためのリスト
	 */
	private List<PackageDoc> packages;

	/**
	 * ドキュメントを生成します。
	 *
	 * @param rootDoc
	 *            Javadoc のルートドキュメント
	 * @throws IOException
	 */
	public void create(RootDoc rootDoc) throws IOException {

		// 例外捕獲
		try {

			// Javadoc のルートドキュメントを取得
			root = rootDoc;

			// Word 文書を生成
			word = new XWPFDocument();

			// ヘッダとフッタを作成
			makeHeaderFooter(Options.getOption("title") + " " + Options.getOption("subtitle"), true);
			makeHeaderFooter(Options.getOption("copyright"), false);

			// 表紙を作成
			makeCoverPage();

			// 出力済パッケージリストを初期化
			packages = new ArrayList<PackageDoc>();

			// 全てのクラスを出力
			makeClassPages();

			// Word ファイル保存
			word.write(new FileOutputStream(Options.getOption("file", "document.docx")));

		} finally {

			// Word 文書を閉じる
			if (word != null) {
				try {
					(word).close();
				} catch (Exception e) {
				}
			}
		}
	}

	/**
	 * 表紙を作成します。
	 */
	private void makeCoverPage() {

		// POI 操作
		XWPFRun run;

		// 日付を取得
		Locale locale = new Locale("ja", "JP", "JP");
		Calendar cal = Calendar.getInstance(locale);
		DateFormat jformat = new SimpleDateFormat("GGGGy年M月d日", locale);
		String stamp = jformat.format(cal.getTime());

		// 表紙の情報を出力
		run = DocumentStyle.setCoverParagraph(word.createParagraph(), 800);
		run.setFontSize(28);
		run.setBold(true);
		run.setText(Options.getOption("title"));
		run = DocumentStyle.setCoverParagraph(word.createParagraph(), 200);
		run.setFontSize(20);
		run.setBold(true);
		run.setText(Options.getOption("subtitle"));
		run = DocumentStyle.setCoverParagraph(word.createParagraph(), 300);
		run.setFontSize(18);
		run.setText(Options.getOption("version"));
		run = DocumentStyle.setCoverParagraph(word.createParagraph(), 800);
		run.setFontSize(16);
		run.setText(stamp);
		run = DocumentStyle.setCoverParagraph(word.createParagraph(), 300);
		run.setFontSize(20);
		run.setText(Options.getOption("company"));
	}

	/**
	 * 改ページを挿入します。
	 */
	private void newPage() {
		word.getLastParagraph().createRun().addBreak(BreakType.PAGE);
	}

	/**
	 * ヘッダとフッタを作成します。
	 *
	 * @param str
	 *            出力する文字列
	 * @param isHeader
	 *            ヘッダの場合は true, フッタの場合は false
	 * @throws IOException
	 */
	private void makeHeaderFooter(String str, boolean isHeader) throws IOException {
		CTSectPr sect = word.getDocument().getBody().addNewSectPr();
		XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(word, sect);
		CTP ctp = CTP.Factory.newInstance();
		XWPFParagraph paragraph = new XWPFParagraph(ctp, word);
		XWPFRun run = paragraph.createRun();
		run.setText(str);
		run.setFontFamily("Meiryo UI");
		run.setFontSize(8);
		if (isHeader) {
			paragraph.setAlignment(ParagraphAlignment.LEFT);
			XWPFParagraph[] paragraphs = new XWPFParagraph[] { paragraph };
			policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, paragraphs);
		} else {
			CTP ctp2 = CTP.Factory.newInstance();
			ctp2.addNewR().addNewPgNum();
			XWPFParagraph pagenum = new XWPFParagraph(ctp2, word);
			pagenum.setAlignment(ParagraphAlignment.CENTER);
			paragraph.setAlignment(ParagraphAlignment.RIGHT);
			XWPFParagraph[] paragraphs = new XWPFParagraph[] { pagenum, paragraph };
			policy.createFooter(XWPFHeaderFooterPolicy.DEFAULT, paragraphs);
		}
	}

	/**
	 * 実行メソッドの引数の書式を文字列で取得します。
	 *
	 * @param parameters
	 *            引数の情報
	 * @return 引数の書式を示した文字列
	 */
	private String getParamSignature(Parameter[] parameters) {
		StringBuilder sb = new StringBuilder();
		for (Parameter parameter : parameters) {
			if (0 < sb.length()) {
				sb.append(", ");
			}
			String type = parameter.type().toString();
			type = type.replaceAll("java\\.(lang|util|io|nio)\\.", "");
			sb.append(type);
			sb.append(" ");
			sb.append(parameter.name());
		}
		return sb.toString();
	}

	/**
	 * パラメタに設定されたコメントを取得します。
	 *
	 * @param tags
	 *            タグ情報
	 * @param name
	 *            パラメタ名
	 * @return コメント情報
	 */
	private String getParamComment(ParamTag[] tags, String name) {
		for (ParamTag tag : tags) {
			if (tag.parameterName().equals(name)) {
				return tag.parameterComment();
			}
		}
		return "";
	}

	/**
	 * 例外に設定されたコメントを取得します。
	 *
	 * @param tags
	 *            タグ情報
	 * @param name
	 *            例外クラスの名前
	 * @return コメント情報
	 */
	private String getThrowsComment(ThrowsTag[] tags, String name) {
		for (ThrowsTag tag : tags) {
			if (tag.exceptionName().equals(name)) {
				return tag.exceptionComment();
			}
		}
		return "";
	}

	/**
	 * 全てのクラスの情報を出力します。
	 */
	private void makeClassPages() {

		// 出力文字
		String str;

		// 全てのクラス
		for (ClassDoc classDoc : root.classes()) {

			// POI 操作
			XWPFRun run;

			// パッケージ
			PackageDoc packageDoc = classDoc.containingPackage();

			// 新たなパッケージの場合
			if (!packages.contains(packageDoc)) {

				// 改ページ
				newPage();

				// パッケージ名
				run = DocumentStyle.setChapterTitleParagraph(word.createParagraph(), 0);
				print(run, packageDoc.name() + " パッケージ");

				// パッケージ説明
				str = packageDoc.commentText();
				if (!str.isEmpty()) {
					DocumentStyle.setSeparatorParagraph(word.createParagraph());
					run = DocumentStyle.getDefaultRun(word.createParagraph(), 0);
					print(run, str);
				}

				// 出力済パッケージに追加
				packages.add(packageDoc);
			}

			// 改ページ
			newPage();

			// パッケージ名
			run = DocumentStyle.getDefaultRun(word.createParagraph(), 0);
			print(run, classDoc.containingPackage().name() + " パッケージ");

			// クラス
			run = DocumentStyle.setChapterTitleParagraph(word.createParagraph(), 100);
			print(run, classDoc.name() + " クラス");

			// 継承階層
			List<ClassDoc> classDocs = new ArrayList<ClassDoc>();
			classDocs.add(classDoc);
			ClassDoc d = classDoc.superclass();
			while (d != null) {
				classDocs.add(d);
				d = d.superclass();
			}
			run = DocumentStyle.getDefaultRun(word.createParagraph(), 0);
			Collections.reverse(classDocs);
			for (int i = 0; i < classDocs.size(); i++) {
				if (0 < i) {
					run.addBreak();
				}
				str = "";
				for (int j = 1; j < i; j++) {
					str += "　　 ";
				}
				if (0 < i) {
					str += "　└ ";
				}
				str += classDocs.get(i).qualifiedName();
				print(run, str);
			}

			// インターフェイス
			if (0 < classDoc.interfaces().length) {
				run = DocumentStyle.setSectionParagraph(word.createParagraph(), 100);
				print(run, "すべての実装されたインタフェース:");
				str = "";
				for (int i = 0; i < classDoc.interfaces().length; i++) {
					if (0 < i) {
						str += ", ";
					}
					str += classDoc.interfaces()[i].qualifiedName();
				}
				run = DocumentStyle.getDefaultRun(word.createParagraph(), 200);
				print(run, str);
			}

			// クラス説明
			run = DocumentStyle.setSubTitleParagraph(word.createParagraph(), 200);
			print(run, classDoc.modifiers() + " " + classDoc.name());
			run = DocumentStyle.getDefaultRun(word.createParagraph(), 0);
			print(run, classDoc.commentText());

			// バージョン
			Tag[] versionTags = classDoc.tags("version");
			if (0 < versionTags.length) {
				run = DocumentStyle.setSectionParagraph(word.createParagraph(), 100);
				print(run, "バージョン:");
				run = DocumentStyle.getDefaultRun(word.createParagraph(), 200);
				for (int i = 0; i < versionTags.length; i++) {
					if (0 < i) {
						run.addBreak();
					}
					print(run, versionTags[i].text());
				}
			}

			// 作成者
			Tag[] authorTags = classDoc.tags("author");
			if (0 < authorTags.length) {
				run = DocumentStyle.setSectionParagraph(word.createParagraph(), 100);
				print(run, "作成者:");
				run = DocumentStyle.getDefaultRun(word.createParagraph(), 200);
				for (int i = 0; i < authorTags.length; i++) {
					if (0 < i) {
						run.addBreak();
					}
					print(run, authorTags[i].text());
				}
			}

			// 全ての定数
			if (0 < classDoc.enumConstants().length) {
				run = DocumentStyle.setTitleParagraph(word.createParagraph(), 100);
				print(run, "定数の詳細");
				for (int i = 0; i < classDoc.enumConstants().length; i++) {
					if (0 < i) {
						DocumentStyle.setSeparatorParagraph(word.createParagraph());
					}
					writeFieldDoc(classDoc.enumConstants()[i]);
				}
			}

			// 全てのフィールド
			if (0 < classDoc.fields().length) {
				run = DocumentStyle.setTitleParagraph(word.createParagraph(), 100);
				print(run, "フィールドの詳細");
				for (int i = 0; i < classDoc.fields().length; i++) {
					if (0 < i) {
						DocumentStyle.setSeparatorParagraph(word.createParagraph());
					}
					writeFieldDoc(classDoc.fields()[i]);
				}
			}

			// 全てのコンストラクタ
			if (0 < classDoc.constructors().length) {
				run = DocumentStyle.setTitleParagraph(word.createParagraph(), 100);
				print(run, "コンストラクタの詳細");
				for (int i = 0; i < classDoc.constructors().length; i++) {
					if (0 < i) {
						DocumentStyle.setSeparatorParagraph(word.createParagraph());
					}
					writeMemberDoc(classDoc.constructors()[i]);
				}
			}

			// 全てのメソッド
			if (0 < classDoc.methods().length) {
				run = DocumentStyle.setTitleParagraph(word.createParagraph(), 100);
				print(run, "メソッドの詳細");
				for (int i = 0; i < classDoc.methods().length; i++) {
					if (0 < i) {
						DocumentStyle.setSeparatorParagraph(word.createParagraph());
					}
					writeMemberDoc(classDoc.methods()[i]);
				}
			}
		}
	}

	/**
	 * 全てのフィールドの情報を出力します。
	 *
	 * @param doc
	 *            メンバ情報
	 */
	private void writeFieldDoc(MemberDoc doc) {

		// 種類名
		String fieldType;
		if (doc.isEnumConstant()) {
			fieldType = "列挙型定数";
		} else if (doc.isEnum()) {
			fieldType = "列挙型";
		} else {
			fieldType = "フィールド";
		}

		// フィールド情報
		XWPFRun run;
		run = DocumentStyle.setSubTitleParagraph(word.createParagraph(), 100);
		print(run, doc.name() + " " + fieldType);
		run = DocumentStyle.getDefaultRun(word.createParagraph(), 0);
		print(run, doc.modifiers() + " " + doc.name());
		run = DocumentStyle.getDefaultRun(word.createParagraph(), 200);
		print(run, doc.commentText());
	}

	/**
	 * 全ての実行可能メンバの情報を出力します。
	 *
	 * @param doc
	 *            実行可能メンバの情報
	 */
	private void writeMemberDoc(ExecutableMemberDoc doc) {

		// 出力文字
		String str;

		// 種類名
		String memberType;
		if (doc.isConstructor()) {
			memberType = "コンストラクタ";
		} else if (doc.isMethod()) {
			memberType = "メソッド";
		} else {
			memberType = "メンバ";
		}

		// メソッド情報
		XWPFRun run;
		run = DocumentStyle.setSubTitleParagraph(word.createParagraph(), 100);
		print(run, doc.name() + " " + memberType);
		run = DocumentStyle.getDefaultRun(word.createParagraph(), 0);
		str = doc.modifiers();
		if (doc instanceof MethodDoc) {
			MethodDoc method = (MethodDoc) doc;
			str += " " + method.returnType().simpleTypeName();
		}
		str += " " + doc.name();
		str += " (" + getParamSignature(doc.parameters()) + ")";
		print(run, str);
		if (!doc.commentText().isEmpty()) {
			run = DocumentStyle.getDefaultRun(word.createParagraph(), 200);
			print(run, doc.commentText());
		}

		// パラメータ
		Parameter[] parameters = doc.parameters();
		if (0 < parameters.length) {
			run = DocumentStyle.setSectionParagraph(word.createParagraph(), 100);
			print(run, "パラメータ:");
			for (int i = 0; i < parameters.length; i++) {
				str = String.format("%d) ", i + 1) + parameters[i].name();
				String comment = getParamComment(doc.paramTags(), parameters[i].name());
				if (!comment.isEmpty()) {
					str += " - " + comment;
				}
				run = DocumentStyle.getDefaultRun(word.createParagraph(), 200);
				print(run, str);
			}
		}

		// 戻り値
		if (doc instanceof MethodDoc) {
			MethodDoc method = (MethodDoc) doc;
			if (!method.returnType().simpleTypeName().equals("void")) {
				run = DocumentStyle.setSectionParagraph(word.createParagraph(), 100);
				print(run, "戻り値:");
				str = method.returnType().simpleTypeName();
				Tag[] tags = method.tags("return");
				if (0 < tags.length) {
					String comment = tags[0].text();
					if (!comment.isEmpty()) {
						str += " - " + comment;
					}
				}
				run = DocumentStyle.getDefaultRun(word.createParagraph(), 200);
				print(run, str);
			}
		}

		// 例外
		Type[] exceptions = doc.thrownExceptionTypes();
		if (0 < exceptions.length) {
			run = DocumentStyle.setSectionParagraph(word.createParagraph(), 100);
			print(run, "例外:");
			for (int i = 0; i < exceptions.length; i++) {
				str = exceptions[i].simpleTypeName();
				String comment = getThrowsComment(doc.throwsTags(), exceptions[i].typeName());
				if (!comment.isEmpty()) {
					str += " - " + comment;
				}
				run = DocumentStyle.getDefaultRun(word.createParagraph(), 200);
				print(run, str);
			}
		}
	}

	/**
	 * Javadoc の情報を出力します。
	 * <p>
	 * HTMLタグは簡易的に解釈します。処理しないHTMLタグは削除して文字情報のみ出力します。
	 * <p>
	 * Javadocのインラインタグはフォントを切り替えて文字部分のみ出力します。
	 *
	 * @param run
	 *            文字出力用のハンドル
	 * @param str
	 *            出力する Javadoc 文字情報
	 */
	private void print(XWPFRun run, String str) {

		// 段落ごとに処理
		String[] paragraphs = str.split("\\s*<(p|P)>\\s*");
		for (int i = 0; i < paragraphs.length; i++) {
			if (0 < i) {
				int indent = word.getLastParagraph().getIndentFromLeft();
				run = DocumentStyle.getDefaultRun(word.createParagraph(), indent);
			}

			// 改行の結合
			paragraphs[i] = paragraphs[i].replaceAll("\\s*[\\r\\n]+\\s*", " ");

			// 行ごとに改行を挿入
			paragraphs[i] = paragraphs[i].replaceAll("\\.\\s+", ".\n");
			paragraphs[i] = paragraphs[i].replaceAll("。\\s*", "。\n");

			// 改行ごとに処理
			String[] lines = paragraphs[i].split("\n");
			for (int j = 0; j < lines.length; j++) {
				if (0 < j) {
					run.addCarriageReturn();
				}
				String line = lines[j];

				// HTMLタグの除去
				line = line.replaceAll("\\s*</?([a-z]+|[A-Z]+)>\\s*", "");

				// エンティティ参照の復元
				line = line.replaceAll("&lt;", "<");
				line = line.replaceAll("&gt;", ">");
				line = line.replaceAll("&quot;", "\"");
				line = line.replaceAll("&apos;", "'");
				line = line.replaceAll("&nbsp;", " ");
				line = line.replaceAll("&amp;", "&");

				// Javadocインラインタグの除去(フォント切り替え)
				Pattern p = Pattern.compile("\\{@([a-z]+)\\s*([^\\}]*)\\}");
				Matcher m = p.matcher(line);
				int pos = 0;
				while (m.find()) {
					run.setText(line.substring(pos, m.start()));
					pos = m.end();
					String value = m.group(2).trim();

					// Javadocインラインタグ付き文字として出力
					if (!value.isEmpty()) {
						XWPFRun runTaggedString = DocumentStyle.getDefaultRun(word.getLastParagraph(), -1);
						runTaggedString.setFontFamily(Options.getOption("font2", "Consolas"));
						runTaggedString.setText(value);
						run = DocumentStyle.getDefaultRun(word.getLastParagraph(), -1);
					}
				}

				// 通常文字として出力
				run.setText(line.substring(pos));
			}
		}
	}
}