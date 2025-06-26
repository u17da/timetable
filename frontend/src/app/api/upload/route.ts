import { NextRequest, NextResponse } from 'next/server';
import OpenAI from 'openai';
import * as XLSX from 'xlsx';
import sharp from 'sharp';

interface TimetableData {
  title: string;
  schedule: {
    [key: string]: Array<{
      time: string;
      subject: string;
      room: string;
      normalizedSubject?: string;
      subjectColor?: string;
      isUnmatched?: boolean;
      originalSubject?: string;
    }>;
  };
}

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

interface SubjectMaster {
  [schoolLevel: string]: {
    [grade: string]: {
      [subject: string]: {
        aliases: string[];
        color: string;
      };
    };
  };
}

const EMBEDDED_SUBJECT_MASTER: SubjectMaster = {
  "elementary": {
    "1": {
      "国語": {
        "aliases": ["国語", "こくご"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "書写": {
        "aliases": ["書写", "しょしゃ"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "算数": {
        "aliases": ["算数", "さんすう"],
        "color": "#E1F7FD 算数/数学"
      },
      "音楽": {
        "aliases": ["音楽", "おんがく"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "図工": {
        "aliases": ["図工", "ずこう"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "体育": {
        "aliases": ["体育", "たいいく"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "道徳": {
        "aliases": ["道徳", "どうとく"],
        "color": "#EDE7F6 道徳/生活"
      },
      "生活": {
        "aliases": ["生活", "せいかつ"],
        "color": "#EDE7F6 道徳/生活"
      },
      "特活": {
        "aliases": ["特活", "とっかつ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "学活": {
        "aliases": ["学活", "がっかつ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "行事": {
        "aliases": ["行事", "ぎょうじ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "自立": {
        "aliases": ["自立", "じりつ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "その他": {
        "aliases": ["その他", "そのた"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      }
    },
    "2": {
      "国語": {
        "aliases": ["国語", "こくご"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "書写": {
        "aliases": ["書写", "しょしゃ"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "算数": {
        "aliases": ["算数", "さんすう"],
        "color": "#E1F7FD 算数/数学"
      },
      "音楽": {
        "aliases": ["音楽", "おんがく"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "図工": {
        "aliases": ["図工", "ずこう"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "体育": {
        "aliases": ["体育", "たいいく"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "道徳": {
        "aliases": ["道徳", "どうとく"],
        "color": "#EDE7F6 道徳/生活"
      },
      "生活": {
        "aliases": ["生活", "せいかつ"],
        "color": "#EDE7F6 道徳/生活"
      },
      "特活": {
        "aliases": ["特活", "とっかつ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "学活": {
        "aliases": ["学活", "がっかつ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "行事": {
        "aliases": ["行事", "ぎょうじ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "自立": {
        "aliases": ["自立", "じりつ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "その他": {
        "aliases": ["その他", "そのた"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      }
    },
    "3": {
      "国語": {
        "aliases": ["国語", "こくご"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "書写": {
        "aliases": ["書写", "しょしゃ"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "社会": {
        "aliases": ["社会", "しゃかい"],
        "color": "#FFF8E1 社会/公民/地理/歴史"
      },
      "算数": {
        "aliases": ["算数", "さんすう"],
        "color": "#E1F7FD 算数/数学"
      },
      "理科": {
        "aliases": ["理科", "りか"],
        "color": "#E8F5E9 理科"
      },
      "音楽": {
        "aliases": ["音楽", "おんがく"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "図工": {
        "aliases": ["図工", "ずこう"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "体育": {
        "aliases": ["体育", "たいいく"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "保健": {
        "aliases": ["保健", "ほけん"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "道徳": {
        "aliases": ["道徳", "どうとく"],
        "color": "#EDE7F6 道徳/生活"
      },
      "外国語活動": {
        "aliases": ["外国語活動", "がいこくごかつどう"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "総合": {
        "aliases": ["総合", "そうごう"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "特活": {
        "aliases": ["特活", "とっかつ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "学活": {
        "aliases": ["学活", "がっかつ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "行事": {
        "aliases": ["行事", "ぎょうじ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "自立": {
        "aliases": ["自立", "じりつ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "その他": {
        "aliases": ["その他", "そのた"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      }
    },
    "4": {
      "国語": {
        "aliases": ["国語", "こくご"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "書写": {
        "aliases": ["書写", "しょしゃ"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "社会": {
        "aliases": ["社会", "しゃかい"],
        "color": "#FFF8E1 社会/公民/地理/歴史"
      },
      "算数": {
        "aliases": ["算数", "さんすう"],
        "color": "#E1F7FD 算数/数学"
      },
      "理科": {
        "aliases": ["理科", "りか"],
        "color": "#E8F5E9 理科"
      },
      "音楽": {
        "aliases": ["音楽", "おんがく"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "図工": {
        "aliases": ["図工", "ずこう"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "体育": {
        "aliases": ["体育", "たいいく"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "保健": {
        "aliases": ["保健", "ほけん"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "道徳": {
        "aliases": ["道徳", "どうとく"],
        "color": "#EDE7F6 道徳/生活"
      },
      "外国語活動": {
        "aliases": ["外国語活動", "がいこくごかつどう"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "英語": {
        "aliases": ["英語", "えいご"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "総合": {
        "aliases": ["総合", "そうごう"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "特活": {
        "aliases": ["特活", "とっかつ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "学活": {
        "aliases": ["学活", "がっかつ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "委員会": {
        "aliases": ["委員会", "いいんかい"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "クラブ": {
        "aliases": ["クラブ", "くらぶ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "行事": {
        "aliases": ["行事", "ぎょうじ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "自立": {
        "aliases": ["自立", "じりつ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "その他": {
        "aliases": ["その他", "そのた"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      }
    },
    "5": {
      "国語": {
        "aliases": ["国語", "こくご"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "書写": {
        "aliases": ["書写", "しょしゃ"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "社会": {
        "aliases": ["社会", "しゃかい"],
        "color": "#FFF8E1 社会/公民/地理/歴史"
      },
      "算数": {
        "aliases": ["算数", "さんすう"],
        "color": "#E1F7FD 算数/数学"
      },
      "理科": {
        "aliases": ["理科", "りか"],
        "color": "#E8F5E9 理科"
      },
      "音楽": {
        "aliases": ["音楽", "おんがく"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "図工": {
        "aliases": ["図工", "ずこう"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "家庭": {
        "aliases": ["家庭", "かてい"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "体育": {
        "aliases": ["体育", "たいいく"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "保健": {
        "aliases": ["保健", "ほけん"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "外国語": {
        "aliases": ["外国語", "がいこくご"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "英語": {
        "aliases": ["英語", "えいご"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "道徳": {
        "aliases": ["道徳", "どうとく"],
        "color": "#EDE7F6 道徳/生活"
      },
      "総合": {
        "aliases": ["総合", "そうごう"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "特活": {
        "aliases": ["特活", "とっかつ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "学活": {
        "aliases": ["学活", "がっかつ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "委員会": {
        "aliases": ["委員会", "いいんかい"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "クラブ": {
        "aliases": ["クラブ", "くらぶ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "行事": {
        "aliases": ["行事", "ぎょうじ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "自立": {
        "aliases": ["自立", "じりつ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "その他": {
        "aliases": ["その他", "そのた"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      }
    },
    "6": {
      "国語": {
        "aliases": ["国語", "こくご"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "書写": {
        "aliases": ["書写", "しょしゃ"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "社会": {
        "aliases": ["社会", "しゃかい"],
        "color": "#FFF8E1 社会/公民/地理/歴史"
      },
      "算数": {
        "aliases": ["算数", "さんすう"],
        "color": "#E1F7FD 算数/数学"
      },
      "理科": {
        "aliases": ["理科", "りか"],
        "color": "#E8F5E9 理科"
      },
      "音楽": {
        "aliases": ["音楽", "おんがく"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "図工": {
        "aliases": ["図工", "ずこう"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "家庭": {
        "aliases": ["家庭", "かてい"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "体育": {
        "aliases": ["体育", "たいいく"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "保健": {
        "aliases": ["保健", "ほけん"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "外国語": {
        "aliases": ["外国語", "がいこくご"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "英語": {
        "aliases": ["英語", "えいご"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "道徳": {
        "aliases": ["道徳", "どうとく"],
        "color": "#EDE7F6 道徳/生活"
      },
      "総合": {
        "aliases": ["総合", "そうごう"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "特活": {
        "aliases": ["特活", "とっかつ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "学活": {
        "aliases": ["学活", "がっかつ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "委員会": {
        "aliases": ["委員会", "いいんかい"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "クラブ": {
        "aliases": ["クラブ", "くらぶ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "行事": {
        "aliases": ["行事", "ぎょうじ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "自立": {
        "aliases": ["自立", "じりつ"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "その他": {
        "aliases": ["その他", "そのた"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      }
    }
  },
  "junior": {
    "1": {
      "国語": {
        "aliases": ["国語"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "社会": {
        "aliases": ["社会"],
        "color": "#FFF8E1 社会/公民/地理/歴史"
      },
      "数学": {
        "aliases": ["数学"],
        "color": "#E1F7FD 算数/数学"
      },
      "理科": {
        "aliases": ["理科"],
        "color": "#E8F5E9 理科"
      },
      "音楽": {
        "aliases": ["音楽"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "美術": {
        "aliases": ["美術"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "保健体育": {
        "aliases": ["保健体育"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "技術": {
        "aliases": ["技術"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "家庭": {
        "aliases": ["家庭"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "英語": {
        "aliases": ["英語"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "道徳": {
        "aliases": ["道徳"],
        "color": "#EDE7F6 道徳/生活"
      },
      "総合": {
        "aliases": ["総合"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "特活": {
        "aliases": ["特活"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "学活": {
        "aliases": ["学活"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      }
    },
    "2": {
      "国語": {
        "aliases": ["国語"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "社会": {
        "aliases": ["社会"],
        "color": "#FFF8E1 社会/公民/地理/歴史"
      },
      "数学": {
        "aliases": ["数学"],
        "color": "#E1F7FD 算数/数学"
      },
      "理科": {
        "aliases": ["理科"],
        "color": "#E8F5E9 理科"
      },
      "音楽": {
        "aliases": ["音楽"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "美術": {
        "aliases": ["美術"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "保健体育": {
        "aliases": ["保健体育"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "技術": {
        "aliases": ["技術"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "家庭": {
        "aliases": ["家庭"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "英語": {
        "aliases": ["英語"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "道徳": {
        "aliases": ["道徳"],
        "color": "#EDE7F6 道徳/生活"
      },
      "総合": {
        "aliases": ["総合"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "特活": {
        "aliases": ["特活"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "学活": {
        "aliases": ["学活"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      }
    },
    "3": {
      "国語": {
        "aliases": ["国語"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "社会": {
        "aliases": ["社会"],
        "color": "#FFF8E1 社会/公民/地理/歴史"
      },
      "数学": {
        "aliases": ["数学"],
        "color": "#E1F7FD 算数/数学"
      },
      "理科": {
        "aliases": ["理科"],
        "color": "#E8F5E9 理科"
      },
      "音楽": {
        "aliases": ["音楽"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "美術": {
        "aliases": ["美術"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "保健体育": {
        "aliases": ["保健体育"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "技術": {
        "aliases": ["技術"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "家庭": {
        "aliases": ["家庭"],
        "color": "#FBE9E7 音楽/家庭/図画工作/体育/美術/保健体育/技術"
      },
      "英語": {
        "aliases": ["英語"],
        "color": "#FCE4EC 国語/英語/外国語活動"
      },
      "道徳": {
        "aliases": ["道徳"],
        "color": "#EDE7F6 道徳/生活"
      },
      "総合": {
        "aliases": ["総合"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "特活": {
        "aliases": ["特活"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      },
      "学活": {
        "aliases": ["学活"],
        "color": "#ECEFF1 総合的な学習の時間/特別活動/学級活動"
      }
    }
  }
};

async function loadSubjectMaster(useJsonFile: boolean = false): Promise<SubjectMaster> {
  if (useJsonFile) {
    try {
      const response = await fetch('/subject_master_full.json');
      if (response.ok) {
        return await response.json();
      }
    } catch {
      console.log('Failed to load JSON file, falling back to embedded data');
    }
  }
  return EMBEDDED_SUBJECT_MASTER;
}

function extractColorHex(colorString: string): string {
  const match = colorString.match(/^#[0-9A-Fa-f]{6}/);
  return match ? match[0] : '#E5E7EB';
}

function normalizeSubject(
  subject: string, 
  schoolLevel: string, 
  grade: string, 
  subjectMaster: SubjectMaster
): { normalizedSubject: string; color: string; isUnmatched: boolean } {
  console.log(`Normalizing subject: "${subject}" for ${schoolLevel} grade ${grade}`);
  
  const gradeData = subjectMaster[schoolLevel]?.[grade];
  console.log('Grade data found:', !!gradeData);
  if (gradeData) {
    console.log('Available subjects in grade data:', Object.keys(gradeData));
  }
  
  if (!gradeData) {
    console.log('No grade data found, returning unmatched');
    return { normalizedSubject: subject, color: '#FFFFFF', isUnmatched: true };
  }

  for (const [canonicalSubject, data] of Object.entries(gradeData)) {
    if (data.aliases.includes(subject)) {
      console.log(`Found exact alias match: "${subject}" -> "${canonicalSubject}"`);
      return {
        normalizedSubject: canonicalSubject,
        color: extractColorHex(data.color),
        isUnmatched: false
      };
    }
  }

  for (const [canonicalSubject, data] of Object.entries(gradeData)) {
    if (canonicalSubject.includes(subject) || subject.includes(canonicalSubject)) {
      console.log(`Found fuzzy match: "${subject}" -> "${canonicalSubject}"`);
      return {
        normalizedSubject: canonicalSubject,
        color: extractColorHex(data.color),
        isUnmatched: false
      };
    }
  }

  console.log(`No match found for subject: "${subject}"`);
  return { normalizedSubject: subject, color: '#FFFFFF', isUnmatched: true };
}

const timetablesStorage: Record<string, unknown> = {};
const normalizationCache: Record<string, { id: string; data: TimetableData }> = {};

function getCacheKey(fileContent: string, schoolLevel: string, grade: string): string {
  return `${schoolLevel}_${grade}_${fileContent.slice(0, 100)}`;
}

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get('file') as File;
    const rawGrade = formData.get('grade') as string || '小学1年';
    
    let schoolLevel = 'elementary';
    let grade = '1';
    
    if (rawGrade.startsWith('小学')) {
      schoolLevel = 'elementary';
      grade = rawGrade.replace('小学', '').replace('年', '');
    } else if (rawGrade.startsWith('中学')) {
      schoolLevel = 'junior';
      grade = rawGrade.replace('中学', '').replace('年', '');
    }
    
    if (!file) {
      return NextResponse.json(
        { error: 'No file provided' },
        { status: 400 }
      );
    }

    if (file.type.startsWith('image/')) {
      return await processImageFile(file, schoolLevel, grade);
    } else if (
      file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
      file.type === 'application/vnd.ms-excel'
    ) {
      return await processExcelFile(file, schoolLevel, grade);
    } else {
      return NextResponse.json(
        { error: 'Unsupported file type. Please upload an image or Excel file.' },
        { status: 400 }
      );
    }
  } catch (error) {
    console.error('Upload error:', error);
    return NextResponse.json(
      { error: 'Error processing file' },
      { status: 500 }
    );
  }
}

async function processImageFile(file: File, schoolLevel: string, grade: string) {
  try {
    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    
    const pngBuffer = await sharp(buffer).png().toBuffer();
    const base64Image = pngBuffer.toString('base64');
    
    const cacheKey = getCacheKey(base64Image, schoolLevel, grade);
    if (normalizationCache[cacheKey]) {
      console.log('Using cached normalization result');
      return NextResponse.json(normalizationCache[cacheKey]);
    }

    const subjectMasterForPrompt = await loadSubjectMaster();
    const gradeDataForPrompt = subjectMasterForPrompt[schoolLevel]?.[grade];
    const canonicalSubjects = gradeDataForPrompt ? Object.keys(gradeDataForPrompt) : [];
    const subjectAliases = gradeDataForPrompt ? 
      Object.entries(gradeDataForPrompt).flatMap(([subject, data]) => 
        data.aliases.map(alias => `${alias} → ${subject}`)
      ) : [];
    
    const response = await openai.chat.completions.create({
      model: "gpt-4o",
      temperature: 0,
      messages: [
        {
          role: "user",
          content: [
            {
              type: "text",
              text: `CONTEXT: This is a ${schoolLevel} grade ${grade} timetable.

CANONICAL SUBJECTS for this grade: ${canonicalSubjects.join(', ')}

SUBJECT ALIASES AND MAPPINGS:
${subjectAliases.join('\n')}

NORMALIZATION INSTRUCTIONS:
1. Extract the raw timetable data exactly as it appears
2. For each subject, use your intelligence to match it to the most appropriate canonical subject from the list above
3. When you see abbreviated or alternative forms (like "えいご", "さんすう", "こく語"), map them to their canonical forms using the aliases provided
4. Output the CANONICAL subject names in your JSON response, not the raw extracted text
5. If no good match exists, preserve the original text

Return a JSON object with the following structure:
{
  "title": "Schedule title if visible",
  "schedule": {
    "Monday": [{"time": "09:00-10:00", "subject": "算数", "room": "A101", "originalSubject": "さんすう"}],
    "Tuesday": [{"time": "09:00-10:00", "subject": "国語", "room": "B202", "originalSubject": "こくご"}],
    "Wednesday": [],
    "Thursday": [],
    "Friday": [],
    "Saturday": [],
    "Sunday": []
  }
}

Extract all visible time slots, subjects, and room numbers. Include both the canonical subject name and the original extracted text for reference. Use the canonical subject mappings provided above to ensure accurate normalization.`
            },
            {
              type: "image_url",
              image_url: {
                url: `data:image/png;base64,${base64Image}`
              }
            }
          ]
        }
      ],
      max_tokens: 1000
    });

    const content = response.choices[0].message.content;
    let timetableData;
    
    try {
      timetableData = JSON.parse(content || '{}');
    } catch {
      console.log('OpenAI response parsing failed, attempting fallback extraction');
      const jsonMatch = content?.match(/```(?:json)?\s*(\{[\s\S]*?\})\s*```/);
      if (jsonMatch) {
        try {
          timetableData = JSON.parse(jsonMatch[1]);
        } catch {
          timetableData = {
            title: "Extracted Timetable",
            schedule: {
              Monday: [], Tuesday: [], Wednesday: [], Thursday: [], Friday: [], Saturday: [], Sunday: []
            }
          };
        }
      } else {
        const simpleJsonMatch = content?.match(/\{[\s\S]*\}/);
        if (simpleJsonMatch) {
          try {
            timetableData = JSON.parse(simpleJsonMatch[0]);
          } catch {
            timetableData = {
              title: "Extracted Timetable",
              schedule: {
                Monday: [], Tuesday: [], Wednesday: [], Thursday: [], Friday: [], Saturday: [], Sunday: []
              }
            };
          }
        } else {
          timetableData = {
            title: "Extracted Timetable",
            schedule: {
              Monday: [], Tuesday: [], Wednesday: [], Thursday: [], Friday: [], Saturday: [], Sunday: []
            }
          };
        }
      }
    }

    const subjectMaster = await loadSubjectMaster();
    if (timetableData.schedule) {
      for (const [, entries] of Object.entries(timetableData.schedule)) {
        if (Array.isArray(entries)) {
          for (const entry of entries) {
            if (entry.subject) {
              const exactMatch = normalizeSubject(entry.subject, schoolLevel, grade, subjectMaster);
              if (!exactMatch.isUnmatched) {
                entry.normalizedSubject = exactMatch.normalizedSubject;
                entry.subjectColor = exactMatch.color;
                entry.isUnmatched = false;
                entry.originalSubject = entry.originalSubject || entry.subject;
              } else {
                entry.normalizedSubject = entry.subject;
                entry.subjectColor = '#FFFFFF';
                entry.isUnmatched = true;
                entry.originalSubject = entry.originalSubject || entry.subject;
              }
            }
          }
        }
      }
    }

    const fileId = `img_${Object.keys(timetablesStorage).length}`;
    timetablesStorage[fileId] = timetableData;
    normalizationCache[cacheKey] = {
      id: fileId,
      data: timetableData
    };

    return NextResponse.json({
      id: fileId,
      data: timetableData
    });

  } catch (error) {
    console.error('Image processing error:', error);
    return NextResponse.json(
      { error: 'Error processing image file' },
      { status: 500 }
    );
  }
}

async function processExcelFile(file: File, schoolLevel: string, grade: string) {
  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    const excelText = (jsonData as unknown[][])
      .filter((row) => row.some((cell) => cell !== null && cell !== undefined))
      .map((row) => row.map((cell) => cell?.toString() || '').join('\t'))
      .join('\n');

    const cacheKey = getCacheKey(excelText, schoolLevel, grade);
    if (normalizationCache[cacheKey]) {
      console.log('Using cached normalization result for Excel');
      return NextResponse.json(normalizationCache[cacheKey]);
    }

    const subjectMasterForExcel = await loadSubjectMaster();
    const gradeDataForExcel = subjectMasterForExcel[schoolLevel]?.[grade];
    const canonicalSubjects = gradeDataForExcel ? Object.keys(gradeDataForExcel) : [];
    const subjectAliases = gradeDataForExcel ? 
      Object.entries(gradeDataForExcel).flatMap(([subject, data]) => 
        data.aliases.map(alias => `${alias} → ${subject}`)
      ) : [];

    const response = await openai.chat.completions.create({
      model: "gpt-4o",
      temperature: 0,
      messages: [
        {
          role: "user",
          content: `CONTEXT: This is a ${schoolLevel} grade ${grade} timetable.

CANONICAL SUBJECTS for this grade: ${canonicalSubjects.join(', ')}

SUBJECT ALIASES AND MAPPINGS:
${subjectAliases.join('\n')}

NORMALIZATION INSTRUCTIONS:
1. Extract the raw timetable data exactly as it appears
2. For each subject, use your intelligence to match it to the most appropriate canonical subject from the list above
3. When you see abbreviated or alternative forms (like "えいご", "さんすう", "こく語"), map them to their canonical forms using the aliases provided
4. Output the CANONICAL subject names in your JSON response, not the raw extracted text
5. If no good match exists, preserve the original text

The Excel data is:

${excelText}

Return a JSON object with the following structure:
{
  "title": "Schedule title if identifiable",
  "schedule": {
    "Monday": [{"time": "09:00-10:00", "subject": "算数", "room": "A101", "originalSubject": "さんすう"}],
    "Tuesday": [{"time": "09:00-10:00", "subject": "国語", "room": "B202", "originalSubject": "こくご"}],
    "Wednesday": [],
    "Thursday": [],
    "Friday": [],
    "Saturday": [],
    "Sunday": []
  }
}

Extract all time slots, subjects, and room information. Include both the canonical subject name and the original extracted text for reference. Use the canonical subject mappings provided above to ensure accurate normalization. Organize by weekdays.`
        }
      ],
      max_tokens: 1000
    });

    const content = response.choices[0].message.content;
    let timetableData;
    
    try {
      timetableData = JSON.parse(content || '{}');
    } catch {
      console.log('OpenAI response parsing failed, attempting fallback extraction');
      const jsonMatch = content?.match(/```(?:json)?\s*(\{[\s\S]*?\})\s*```/);
      if (jsonMatch) {
        try {
          timetableData = JSON.parse(jsonMatch[1]);
        } catch {
          timetableData = {
            title: "Extracted Timetable",
            schedule: {
              Monday: [], Tuesday: [], Wednesday: [], Thursday: [], Friday: [], Saturday: [], Sunday: []
            }
          };
        }
      } else {
        const simpleJsonMatch = content?.match(/\{[\s\S]*\}/);
        if (simpleJsonMatch) {
          try {
            timetableData = JSON.parse(simpleJsonMatch[0]);
          } catch {
            timetableData = {
              title: "Extracted Timetable",
              schedule: {
                Monday: [], Tuesday: [], Wednesday: [], Thursday: [], Friday: [], Saturday: [], Sunday: []
              }
            };
          }
        } else {
          timetableData = {
            title: "Extracted Timetable",
            schedule: {
              Monday: [], Tuesday: [], Wednesday: [], Thursday: [], Friday: [], Saturday: [], Sunday: []
            }
          };
        }
      }
    }

    const subjectMaster = await loadSubjectMaster();
    if (timetableData.schedule) {
      for (const [, entries] of Object.entries(timetableData.schedule)) {
        if (Array.isArray(entries)) {
          for (const entry of entries) {
            if (entry.subject) {
              const exactMatch = normalizeSubject(entry.subject, schoolLevel, grade, subjectMaster);
              if (!exactMatch.isUnmatched) {
                entry.normalizedSubject = exactMatch.normalizedSubject;
                entry.subjectColor = exactMatch.color;
                entry.isUnmatched = false;
                entry.originalSubject = entry.originalSubject || entry.subject;
              } else {
                entry.normalizedSubject = entry.subject;
                entry.subjectColor = '#FFFFFF';
                entry.isUnmatched = true;
                entry.originalSubject = entry.originalSubject || entry.subject;
              }
            }
          }
        }
      }
    }

    const fileId = `excel_${Object.keys(timetablesStorage).length}`;
    timetablesStorage[fileId] = timetableData;
    normalizationCache[cacheKey] = {
      id: fileId,
      data: timetableData
    };

    return NextResponse.json({
      id: fileId,
      data: timetableData
    });

  } catch (error) {
    console.error('Excel processing error:', error);
    return NextResponse.json(
      { error: 'Error processing Excel file' },
      { status: 500 }
    );
  }
}
