// Target: .NET Framework 4.8
// Revit API: Autodesk.Revit.UI
// File: RandomVerseCommand.cs
// Purpose: Show a random inspirational Bible verse with its category (Love, Faith, etc.) in a TaskDialog.
// Notes: No external files required. Safe with your hotloader.

using System;
using System.Collections.Generic;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;

namespace ABMEP.Work
{
    [Transaction(TransactionMode.Manual)]
    public class RandomVerseCommand : IExternalCommand
    {
        private struct Verse
        {
            public string Category { get; }
            public string Ref { get; }
            public string Text { get; }
            public Verse(string category, string reference, string text)
            {
                Category = category;
                Ref = reference;
                Text = text;
            }
        }

        private static readonly Random _rng = new Random();

        private static readonly List<Verse> Verses = new List<Verse>
        {
            // =========================
            // LOVE
            // =========================
            new Verse("Love", "John 3:16", "For God so loved the world that he gave his one and only Son, that whoever believes in him shall not perish but have eternal life."),
            new Verse("Love", "1 Corinthians 13:4–7", "Love is patient, love is kind. It does not envy, it does not boast, it is not proud. It does not dishonor others, it is not self-seeking, it is not easily angered, it keeps no record of wrongs. Love does not delight in evil but rejoices with the truth. It always protects, always trusts, always hopes, always perseveres."),
            new Verse("Love", "1 John 4:8", "Whoever does not love does not know God, because God is love."),
            new Verse("Love", "1 John 4:19", "We love because he first loved us."),
            new Verse("Love", "Romans 5:8", "But God demonstrates his own love for us in this: While we were still sinners, Christ died for us."),
            new Verse("Love", "Ephesians 5:25", "Husbands, love your wives, just as Christ loved the church and gave himself up for her."),
            new Verse("Love", "1 Peter 4:8", "Above all, love each other deeply, because love covers over a multitude of sins."),
            new Verse("Love", "Matthew 22:37–39", "Jesus replied: ‘Love the Lord your God with all your heart and with all your soul and with all your mind.’ This is the first and greatest commandment. And the second is like it: ‘Love your neighbor as yourself.’"),
            new Verse("Love", "Proverbs 17:17", "A friend loves at all times, and a brother is born for a time of adversity."),
            new Verse("Love", "Song of Solomon 8:7", "Many waters cannot quench love; rivers cannot sweep it away. If one were to give all the wealth of one’s house for love, it would be utterly scorned."),

            // =========================
            // FAITH
            // =========================
            new Verse("Faith", "Hebrews 11:1", "Now faith is confidence in what we hope for and assurance about what we do not see."),
            new Verse("Faith", "Matthew 17:20", "He replied, ‘Because you have so little faith. Truly I tell you, if you have faith as small as a mustard seed, you can say to this mountain, “Move from here to there,” and it will move. Nothing will be impossible for you.’"),
            new Verse("Faith", "Mark 9:23", "‘If you can’?” said Jesus. “Everything is possible for one who believes.”"),
            new Verse("Faith", "Romans 10:17", "Consequently, faith comes from hearing the message, and the message is heard through the word about Christ."),
            new Verse("Faith", "2 Corinthians 5:7", "For we live by faith, not by sight."),
            new Verse("Faith", "James 1:6", "But when you ask, you must believe and not doubt, because the one who doubts is like a wave of the sea, blown and tossed by the wind."),
            new Verse("Faith", "Matthew 21:22", "If you believe, you will receive whatever you ask for in prayer."),
            new Verse("Faith", "Romans 4:20–21", "Yet he did not waver through unbelief regarding the promise of God, but was strengthened in his faith and gave glory to God, being fully persuaded that God had power to do what he had promised."),
            new Verse("Faith", "Luke 17:5–6", "The apostles said to the Lord, ‘Increase our faith!’ He replied, ‘If you have faith as small as a mustard seed, you can say to this mulberry tree, “Be uprooted and planted in the sea,” and it will obey you.’"),
            new Verse("Faith", "Ephesians 2:8–9", "For it is by grace you have been saved, through faith—and this is not from yourselves, it is the gift of God—not by works, so that no one can boast."),

            // =========================
            // HOPE
            // =========================
            new Verse("Hope", "Jeremiah 29:11", "‘For I know the plans I have for you,’ declares the Lord, ‘plans to prosper you and not to harm you, plans to give you hope and a future.’"),
            new Verse("Hope", "Romans 15:13", "May the God of hope fill you with all joy and peace as you trust in him, so that you may overflow with hope by the power of the Holy Spirit."),
            new Verse("Hope", "Psalm 42:11", "Why, my soul, are you downcast? Why so disturbed within me? Put your hope in God, for I will yet praise him, my Savior and my God."),
            new Verse("Hope", "1 Peter 1:3", "Praise be to the God and Father of our Lord Jesus Christ! In his great mercy he has given us new birth into a living hope through the resurrection of Jesus Christ from the dead."),
            new Verse("Hope", "Romans 5:5", "And hope does not put us to shame, because God’s love has been poured out into our hearts through the Holy Spirit, who has been given to us."),
            new Verse("Hope", "Proverbs 23:18", "There is surely a future hope for you, and your hope will not be cut off."),
            new Verse("Hope", "Isaiah 40:31", "But those who hope in the Lord will renew their strength. They will soar on wings like eagles; they will run and not grow weary, they will walk and not be faint."),
            new Verse("Hope", "Titus 2:13", "While we wait for the blessed hope—the appearing of the glory of our great God and Savior, Jesus Christ."),
            new Verse("Hope", "Psalm 71:5", "For you have been my hope, Sovereign Lord, my confidence since my youth."),
            new Verse("Hope", "Hebrews 6:19", "We have this hope as an anchor for the soul, firm and secure. It enters the inner sanctuary behind the curtain."),

            // =========================
            // PEACE
            // =========================
            new Verse("Peace", "John 14:27", "Peace I leave with you; my peace I give you. I do not give to you as the world gives. Do not let your hearts be troubled and do not be afraid."),
            new Verse("Peace", "Philippians 4:6–7", "Do not be anxious about anything, but in every situation, by prayer and petition, with thanksgiving, present your requests to God. And the peace of God, which transcends all understanding, will guard your hearts and your minds in Christ Jesus."),
            new Verse("Peace", "Isaiah 26:3", "You will keep in perfect peace those whose minds are steadfast, because they trust in you."),
            new Verse("Peace", "Romans 5:1", "Therefore, since we have been justified through faith, we have peace with God through our Lord Jesus Christ."),
            new Verse("Peace", "Colossians 3:15", "Let the peace of Christ rule in your hearts, since as members of one body you were called to peace. And be thankful."),
            new Verse("Peace", "John 16:33", "I have told you these things, so that in me you may have peace. In this world you will have trouble. But take heart! I have overcome the world."),
            new Verse("Peace", "Ephesians 2:14", "For he himself is our peace, who has made the two groups one and has destroyed the barrier, the dividing wall of hostility."),
            new Verse("Peace", "2 Thessalonians 3:16", "Now may the Lord of peace himself give you peace at all times and in every way. The Lord be with all of you."),
            new Verse("Peace", "Matthew 5:9", "Blessed are the peacemakers, for they will be called children of God."),
            new Verse("Peace", "Numbers 6:24–26", "The Lord bless you and keep you; the Lord make his face shine on you and be gracious to you; the Lord turn his face toward you and give you peace."),

            // =========================
            // JOY
            // =========================
            new Verse("Joy", "Nehemiah 8:10", "Do not grieve, for the joy of the Lord is your strength."),
            new Verse("Joy", "Psalm 16:11", "You make known to me the path of life; you will fill me with joy in your presence, with eternal pleasures at your right hand."),
            new Verse("Joy", "John 15:11", "I have told you this so that my joy may be in you and that your joy may be complete."),
            new Verse("Joy", "Romans 15:13", "May the God of hope fill you with all joy and peace as you trust in him, so that you may overflow with hope by the power of the Holy Spirit."),
            new Verse("Joy", "Philippians 4:4", "Rejoice in the Lord always. I will say it again: Rejoice!"),
            new Verse("Joy", "Psalm 30:5", "For his anger lasts only a moment, but his favor lasts a lifetime; weeping may stay for the night, but rejoicing comes in the morning."),
            new Verse("Joy", "1 Peter 1:8", "Though you have not seen him, you love him; and even though you do not see him now, you believe in him and are filled with an inexpressible and glorious joy."),
            new Verse("Joy", "James 1:2–3", "Consider it pure joy, my brothers and sisters, whenever you face trials of many kinds, because you know that the testing of your faith produces perseverance."),
            new Verse("Joy", "Isaiah 55:12", "You will go out in joy and be led forth in peace; the mountains and hills will burst into song before you, and all the trees of the field will clap their hands."),
            new Verse("Joy", "Romans 12:12", "Be joyful in hope, patient in affliction, faithful in prayer."),

            // =========================
            // STRENGTH
            // =========================
            new Verse("Strength", "Isaiah 40:31", "But those who hope in the Lord will renew their strength. They will soar on wings like eagles; they will run and not grow weary, they will walk and not be faint."),
            new Verse("Strength", "Philippians 4:13", "I can do all this through him who gives me strength."),
            new Verse("Strength", "Nehemiah 8:10", "Do not grieve, for the joy of the Lord is your strength."),
            new Verse("Strength", "Psalm 18:1–2", "I love you, Lord, my strength. The Lord is my rock, my fortress and my deliverer; my God is my rock, in whom I take refuge, my shield and the horn of my salvation, my stronghold."),
            new Verse("Strength", "Isaiah 41:10", "So do not fear, for I am with you; do not be dismayed, for I am your God. I will strengthen you and help you; I will uphold you with my righteous right hand."),
            new Verse("Strength", "2 Corinthians 12:9–10", "But he said to me, ‘My grace is sufficient for you, for my power is made perfect in weakness.’ Therefore I will boast all the more gladly about my weaknesses, so that Christ’s power may rest on me… For when I am weak, then I am strong."),
            new Verse("Strength", "Ephesians 6:10", "Finally, be strong in the Lord and in his mighty power."),
            new Verse("Strength", "Psalm 29:11", "The Lord gives strength to his people; the Lord blesses his people with peace."),
            new Verse("Strength", "Proverbs 24:10", "If you falter in a time of trouble, how small is your strength!"),
            new Verse("Strength", "Deuteronomy 31:6", "Be strong and courageous. Do not be afraid or terrified because of them, for the Lord your God goes with you; he will never leave you nor forsake you."),

            // =========================
            // GRACE
            // =========================
            new Verse("Grace", "Ephesians 2:8–9", "For it is by grace you have been saved, through faith—and this is not from yourselves, it is the gift of God—not by works, so that no one can boast."),
            new Verse("Grace", "2 Corinthians 12:9", "But he said to me, ‘My grace is sufficient for you, for my power is made perfect in weakness.’"),
            new Verse("Grace", "Romans 5:8", "But God demonstrates his own love for us in this: While we were still sinners, Christ died for us."),
            new Verse("Grace", "Titus 2:11", "For the grace of God has appeared that offers salvation to all people."),
            new Verse("Grace", "2 Timothy 1:9", "He has saved us and called us to a holy life—not because of anything we have done but because of his own purpose and grace. This grace was given us in Christ Jesus before the beginning of time."),
            new Verse("Grace", "John 1:16", "Out of his fullness we have all received grace in place of grace already given."),
            new Verse("Grace", "Romans 6:14", "For sin shall no longer be your master, because you are not under the law, but under grace."),
            new Verse("Grace", "1 Peter 5:10", "And the God of all grace, who called you to his eternal glory in Christ, after you have suffered a little while, will himself restore you and make you strong, firm and steadfast."),
            new Verse("Grace", "Hebrews 4:16", "Let us then approach God’s throne of grace with confidence, so that we may receive mercy and find grace to help us in our time of need."),
            new Verse("Grace", "Acts 4:33", "With great power the apostles continued to testify to the resurrection of the Lord Jesus. And God’s grace was so powerfully at work in them all."),

            // =========================
            // FORGIVENESS
            // =========================
            new Verse("Forgiveness", "Ephesians 4:32", "Be kind and compassionate to one another, forgiving each other, just as in Christ God forgave you."),
            new Verse("Forgiveness", "Matthew 6:14–15", "For if you forgive other people when they sin against you, your heavenly Father will also forgive you. But if you do not forgive others their sins, your Father will not forgive your sins."),
            new Verse("Forgiveness", "Colossians 3:13", "Bear with each other and forgive one another if any of you has a grievance against someone. Forgive as the Lord forgave you."),
            new Verse("Forgiveness", "Mark 11:25", "And when you stand praying, if you hold anything against anyone, forgive them, so that your Father in heaven may forgive you your sins."),
            new Verse("Forgiveness", "1 John 1:9", "If we confess our sins, he is faithful and just and will forgive us our sins and purify us from all unrighteousness."),
            new Verse("Forgiveness", "Luke 6:37", "Do not judge, and you will not be judged. Do not condemn, and you will not be condemned. Forgive, and you will be forgiven."),
            new Verse("Forgiveness", "Matthew 18:21–22", "Then Peter came to Jesus and asked, ‘Lord, how many times shall I forgive my brother or sister who sins against me? Up to seven times?’ Jesus answered, ‘I tell you, not seven times, but seventy-seven times.’"),
            new Verse("Forgiveness", "Psalm 103:12", "As far as the east is from the west, so far has he removed our transgressions from us."),
            new Verse("Forgiveness", "Luke 23:34", "Jesus said, ‘Father, forgive them, for they do not know what they are doing.’"),
            new Verse("Forgiveness", "Isaiah 55:7", "Let the wicked forsake their ways and the unrighteous their thoughts. Let them turn to the Lord, and he will have mercy on them, and to our God, for he will freely pardon."),

            // =========================
            // SALVATION
            // =========================
            new Verse("Salvation", "John 3:16", "For God so loved the world that he gave his one and only Son, that whoever believes in him shall not perish but have eternal life."),
            new Verse("Salvation", "Romans 10:9", "If you declare with your mouth, ‘Jesus is Lord,’ and believe in your heart that God raised him from the dead, you will be saved."),
            new Verse("Salvation", "Ephesians 2:8–9", "For it is by grace you have been saved, through faith—and this is not from yourselves, it is the gift of God—not by works, so that no one can boast."),
            new Verse("Salvation", "Titus 3:5", "He saved us, not because of righteous things we had done, but because of his mercy. He saved us through the washing of rebirth and renewal by the Holy Spirit."),
            new Verse("Salvation", "Acts 4:12", "Salvation is found in no one else, for there is no other name under heaven given to mankind by which we must be saved."),
            new Verse("Salvation", "Romans 5:8", "But God demonstrates his own love for us in this: While we were still sinners, Christ died for us."),
            new Verse("Salvation", "John 14:6", "Jesus answered, ‘I am the way and the truth and the life. No one comes to the Father except through me.’"),
            new Verse("Salvation", "Acts 16:31", "They replied, ‘Believe in the Lord Jesus, and you will be saved—you and your household.’"),
            new Verse("Salvation", "1 Timothy 2:4", "Who wants all people to be saved and to come to a knowledge of the truth."),
            new Verse("Salvation", "Hebrews 9:28", "So Christ was sacrificed once to take away the sins of many; and he will appear a second time, not to bear sin, but to bring salvation to those who are waiting for him."),

            // =========================
            // MERCY
            // =========================
            new Verse("Mercy", "Matthew 5:7", "Blessed are the merciful, for they will be shown mercy."),
            new Verse("Mercy", "Ephesians 2:4–5", "But because of his great love for us, God, who is rich in mercy, made us alive with Christ even when we were dead in transgressions—it is by grace you have been saved."),
            new Verse("Mercy", "Titus 3:5", "He saved us, not because of righteous things we had done, but because of his mercy. He saved us through the washing of rebirth and renewal by the Holy Spirit."),
            new Verse("Mercy", "Psalm 103:8", "The Lord is compassionate and gracious, slow to anger, abounding in love."),
            new Verse("Mercy", "Luke 6:36", "Be merciful, just as your Father is merciful."),
            new Verse("Mercy", "James 2:13", "Because judgment without mercy will be shown to anyone who has not been merciful. Mercy triumphs over judgment."),
            new Verse("Mercy", "Romans 9:15", "For he says to Moses, ‘I will have mercy on whom I have mercy, and I will have compassion on whom I have compassion.’"),
            new Verse("Mercy", "Micah 7:18", "Who is a God like you, who pardons sin and forgives the transgression of the remnant of his inheritance? You do not stay angry forever but delight to show mercy."),
            new Verse("Mercy", "Hebrews 4:16", "Let us then approach God's throne of grace with confidence, so that we may receive mercy and find grace to help us in our time of need."),
            new Verse("Mercy", "2 Corinthians 1:3–4", "Praise be to the God and Father of our Lord Jesus Christ, the Father of compassion and the God of all comfort, who comforts us in all our troubles, so that we can comfort those in any trouble with the comfort we ourselves receive from God."),

            // =========================
            // HEALING
            // =========================
            new Verse("Healing", "Isaiah 53:5", "But he was pierced for our transgressions, he was crushed for our iniquities; the punishment that brought us peace was on him, and by his wounds we are healed."),
            new Verse("Healing", "Jeremiah 30:17", "‘But I will restore you to health and heal your wounds,’ declares the Lord, ‘because you are called an outcast, Zion for whom no one cares.’"),
            new Verse("Healing", "Psalm 147:3", "He heals the brokenhearted and binds up their wounds."),
            new Verse("Healing", "1 Peter 2:24", "He himself bore our sins in his body on the cross, so that we might die to sins and live for righteousness; by his wounds you have been healed."),
            new Verse("Healing", "Exodus 15:26", "He said, ‘If you listen carefully to the Lord your God and do what is right in his eyes, if you pay attention to his commands and keep all his decrees, I will not bring on you any of the diseases I brought on the Egyptians, for I am the Lord, who heals you.’"),
            new Verse("Healing", "Matthew 9:35", "Jesus went through all the towns and villages, teaching in their synagogues, proclaiming the good news of the kingdom and healing every disease and sickness."),
            new Verse("Healing", "James 5:15", "And the prayer offered in faith will make the sick person well; the Lord will raise them up. If they have sinned, they will be forgiven."),
            new Verse("Healing", "Mark 5:34", "He said to her, ‘Daughter, your faith has healed you. Go in peace and be freed from your suffering.’"),
            new Verse("Healing", "Proverbs 4:20–22", "My son, pay attention to what I say; turn your ear to my words. Do not let them out of your sight, keep them within your heart; for they are life to those who find them and health to one’s whole body."),
            new Verse("Healing", "Luke 6:19", "And the people all tried to touch him, because power was coming from him and healing them all."),

            // =========================
            // TRUST
            // =========================
            new Verse("Trust", "Proverbs 3:5–6", "Trust in the Lord with all your heart and lean not on your own understanding; in all your ways submit to him, and he will make your paths straight."),
            new Verse("Trust", "Jeremiah 17:7–8", "But blessed is the one who trusts in the Lord, whose confidence is in him. They will be like a tree planted by the water that sends out its roots by the stream. It does not fear when heat comes; its leaves are always green. It has no worries in a year of drought and never fails to bear fruit."),
            new Verse("Trust", "Psalm 9:10", "Those who know your name trust in you, for you, Lord, have never forsaken those who seek you."),
            new Verse("Trust", "Isaiah 26:3", "You will keep in perfect peace those whose minds are steadfast, because they trust in you."),
            new Verse("Trust", "Psalm 37:5", "Commit your way to the Lord; trust in him and he will do this."),
            new Verse("Trust", "Proverbs 16:3", "Commit to the Lord whatever you do, and he will establish your plans."),
            new Verse("Trust", "Psalm 56:3", "When I am afraid, I put my trust in you."),
            new Verse("Trust", "2 Samuel 22:31", "As for God, his way is perfect: The Lord’s word is flawless; he shields all who take refuge in him."),
            new Verse("Trust", "Matthew 6:25–26", "Therefore I tell you, do not worry about your life, what you will eat or drink; or about your body, what you will wear. Is not life more than food, and the body more than clothes? Look at the birds of the air; they do not sow or reap or store away in barns, and yet your heavenly Father feeds them. Are you not much more valuable than they?"),
            new Verse("Trust", "Romans 15:13", "May the God of hope fill you with all joy and peace as you trust in him, so that you may overflow with hope by the power of the Holy Spirit."),

            // =========================
            // PATIENCE
            // =========================
            new Verse("Patience", "James 5:7–8", "Be patient, then, brothers and sisters, until the Lord’s coming. See how the farmer waits for the land to yield its valuable crop, patiently waiting for the autumn and spring rains. You too, be patient and stand firm, because the Lord’s coming is near."),
            new Verse("Patience", "Romans 12:12", "Be joyful in hope, patient in affliction, faithful in prayer."),
            new Verse("Patience", "Psalm 37:7", "Be still before the Lord and wait patiently for him; do not fret when people succeed in their ways, when they carry out their wicked schemes."),
            new Verse("Patience", "Galatians 6:9", "Let us not become weary in doing good, for at the proper time we will reap a harvest if we do not give up."),
            new Verse("Patience", "Proverbs 14:29", "Whoever is patient has great understanding, but one who is quick-tempered displays folly."),
            new Verse("Patience", "Lamentations 3:25–26", "The Lord is good to those whose hope is in him, to the one who seeks him; it is good to wait quietly for the salvation of the Lord."),
            new Verse("Patience", "Colossians 3:12", "Therefore, as God’s chosen people, holy and dearly loved, clothe yourselves with compassion, kindness, humility, gentleness and patience."),
            new Verse("Patience", "Hebrews 10:36", "You need to persevere so that when you have done the will of God, you will receive what he has promised."),
            new Verse("Patience", "Ecclesiastes 7:8", "The end of a matter is better than its beginning, and patience is better than pride."),
            new Verse("Patience", "1 Timothy 6:11", "But you, man of God, flee from all this, and pursue righteousness, godliness, faith, love, endurance and gentleness."),

            // =========================
            // WISDOM
            // =========================
            new Verse("Wisdom", "Proverbs 3:13–18", "Blessed are those who find wisdom, those who gain understanding, for she is more profitable than silver and yields better returns than gold. She is more precious than rubies; nothing you desire can compare with her. Long life is in her right hand; in her left hand are riches and honor. Her ways are pleasant ways, and all her paths are peace. She is a tree of life to those who take hold of her; those who hold her fast will be blessed."),
            new Verse("Wisdom", "James 1:5", "If any of you lacks wisdom, you should ask God, who gives generously to all without finding fault, and it will be given to you."),
            new Verse("Wisdom", "Proverbs 4:7", "The beginning of wisdom is this: Get wisdom. Though it cost all you have, get understanding."),
            new Verse("Wisdom", "Proverbs 2:6", "For the Lord gives wisdom; from his mouth come knowledge and understanding."),
            new Verse("Wisdom", "Colossians 2:2–3", "My goal is that they may be encouraged in heart and united in love, so that they may have the full riches of complete understanding, in order that they may know the mystery of God, namely, Christ, in whom are hidden all the treasures of wisdom and knowledge."),
            new Verse("Wisdom", "Proverbs 9:10", "The fear of the Lord is the beginning of wisdom, and knowledge of the Holy One is understanding."),
            new Verse("Wisdom", "Ecclesiastes 7:12", "Wisdom preserves those who have it."),
            new Verse("Wisdom", "Proverbs 16:16", "How much better to get wisdom than gold, to get insight rather than silver!"),
            new Verse("Wisdom", "1 Corinthians 1:25", "For the foolishness of God is wiser than human wisdom, and the weakness of God is stronger than human strength."),
            new Verse("Wisdom", "Proverbs 19:20", "Listen to advice and accept discipline, and at the end you will be counted among the wise."),

            // =========================
            // COURAGE
            // =========================
            new Verse("Courage", "Joshua 1:9", "Have I not commanded you? Be strong and courageous. Do not be afraid; do not be discouraged, for the Lord your God will be with you wherever you go."),
            new Verse("Courage", "Isaiah 41:10", "So do not fear, for I am with you; do not be dismayed, for I am your God. I will strengthen you and help you; I will uphold you with my righteous right hand."),
            new Verse("Courage", "Psalm 27:14", "Wait for the Lord; be strong and take heart and wait for the Lord."),
            new Verse("Courage", "Deuteronomy 31:6", "Be strong and courageous. Do not be afraid or terrified because of them, for the Lord your God goes with you; he will never leave you nor forsake you."),
            new Verse("Courage", "2 Timothy 1:7", "For the Spirit God gave us does not make us timid, but gives us power, love and self-discipline."),
            new Verse("Courage", "Psalm 31:24", "Be strong and take heart, all you who hope in the Lord."),
            new Verse("Courage", "John 16:33", "I have told you these things, so that in me you may have peace. In this world you will have trouble. But take heart! I have overcome the world."),
            new Verse("Courage", "1 Corinthians 16:13", "Be on your guard; stand firm in the faith; be courageous; be strong."),
            new Verse("Courage", "Matthew 14:27", "But Jesus immediately said to them: ‘Take courage! It is I. Don’t be afraid.’"),
            new Verse("Courage", "Ephesians 6:10", "Finally, be strong in the Lord and in his mighty power."),

            // =========================
            // OBEDIENCE
            // =========================
            new Verse("Obedience", "John 14:15", "If you love me, keep my commands."),
            new Verse("Obedience", "1 Samuel 15:22", "But Samuel replied: ‘Does the Lord delight in burnt offerings and sacrifices as much as in obeying the Lord? To obey is better than sacrifice, and to heed is better than the fat of rams.’"),
            new Verse("Obedience", "Deuteronomy 5:33", "Walk in obedience to all that the Lord your God has commanded you, so that you may live and prosper and prolong your days in the land that you will possess."),
            new Verse("Obedience", "Romans 6:16", "Don’t you know that when you offer yourselves to someone as obedient slaves, you are slaves of the one you obey—whether you are slaves to sin, which leads to death, or to obedience, which leads to righteousness?"),
            new Verse("Obedience", "Hebrews 5:9", "And, once made perfect, he became the source of eternal salvation for all who obey him."),
            new Verse("Obedience", "James 1:22", "Do not merely listen to the word, and so deceive yourselves. Do what it says."),
            new Verse("Obedience", "Colossians 3:20", "Children, obey your parents in everything, for this pleases the Lord."),
            new Verse("Obedience", "1 Peter 1:14–15", "As obedient children, do not conform to the evil desires you had when you lived in ignorance. But just as he who called you is holy, so be holy in all you do."),
            new Verse("Obedience", "Matthew 7:24", "Therefore everyone who hears these words of mine and puts them into practice is like a wise man who built his house on the rock."),
            new Verse("Obedience", "Exodus 19:5", "Now if you obey me fully and keep my covenant, then out of all nations you will be my treasured possession. Although the whole earth is mine."),

            // =========================
            // REPENTANCE
            // =========================
            new Verse("Repentance", "Acts 3:19", "Repent, then, and turn to God, so that your sins may be wiped out, that times of refreshing may come from the Lord."),
            new Verse("Repentance", "2 Peter 3:9", "The Lord is not slow in keeping his promise, as some understand slowness. Instead he is patient with you, not wanting anyone to perish, but everyone to come to repentance."),
            new Verse("Repentance", "1 John 1:9", "If we confess our sins, he is faithful and just and will forgive us our sins and purify us from all unrighteousness."),
            new Verse("Repentance", "Matthew 4:17", "From that time on Jesus began to preach, ‘Repent, for the kingdom of heaven has come near.’"),
            new Verse("Repentance", "Luke 5:32", "I have not come to call the righteous, but sinners to repentance."),
            new Verse("Repentance", "Revelation 3:19", "Those whom I love I rebuke and discipline. So be earnest and repent."),
            new Verse("Repentance", "Acts 17:30", "In the past God overlooked such ignorance, but now he commands all people everywhere to repent."),
            new Verse("Repentance", "Joel 2:13", "Rend your heart and not your garments. Return to the Lord your God, for he is gracious and compassionate, slow to anger and abounding in love, and he relents from sending calamity."),
            new Verse("Repentance", "Ezekiel 18:30", "‘Therefore, you Israelites, I will judge each of you according to your ways,’ declares the Sovereign Lord. ‘Repent! Turn away from all your offenses; then sin will not be your downfall.’"),
            new Verse("Repentance", "Luke 15:7", "I tell you that in the same way there will be more rejoicing in heaven over one sinner who repents than over ninety-nine righteous persons who do not need to repent."),

            // =========================
            // PRAYER
            // =========================
            new Verse("Prayer", "1 Thessalonians 5:17", "Pray continually."),
            new Verse("Prayer", "Matthew 7:7", "Ask and it will be given to you; seek and you will find; knock and the door will be opened to you."),
            new Verse("Prayer", "Philippians 4:6", "Do not be anxious about anything, but in every situation, by prayer and petition, with thanksgiving, present your requests to God."),
            new Verse("Prayer", "Mark 11:24", "Therefore I tell you, whatever you ask for in prayer, believe that you have received it, and it will be yours."),
            new Verse("Prayer", "James 5:16", "Therefore confess your sins to each other and pray for each other so that you may be healed. The prayer of a righteous person is powerful and effective."),
            new Verse("Prayer", "Luke 11:9–10", "So I say to you: Ask and it will be given to you; seek and you will find; knock and the door will be opened to you. For everyone who asks receives; the one who seeks finds; and to the one who knocks, the door will be opened."),
            new Verse("Prayer", "Romans 12:12", "Be joyful in hope, patient in affliction, faithful in prayer."),
            new Verse("Prayer", "Matthew 6:6", "But when you pray, go into your room, close the door and pray to your Father, who is unseen. Then your Father, who sees what is done in secret, will reward you."),
            new Verse("Prayer", "Jeremiah 29:12", "Then you will call on me and come and pray to me, and I will listen to you."),
            new Verse("Prayer", "1 John 5:14–15", "This is the confidence we have in approaching God: that if we ask anything according to his will, he hears us. And if we know that he hears us—whatever we ask—we know that we have what we asked of him."),

            // =========================
            // PRAISE
            // =========================
            new Verse("Praise", "Psalm 150:6", "Let everything that has breath praise the Lord. Praise the Lord."),
            new Verse("Praise", "Psalm 34:1", "I will extol the Lord at all times; his praise will always be on my lips."),
            new Verse("Praise", "Hebrews 13:15", "Through Jesus, therefore, let us continually offer to God a sacrifice of praise—the fruit of lips that openly profess his name."),
            new Verse("Praise", "Psalm 100:4", "Enter his gates with thanksgiving and his courts with praise; give thanks to him and praise his name."),
            new Verse("Praise", "Psalm 47:6–7", "Sing praises to God, sing praises; sing praises to our King, sing praises. For God is the King of all the earth; sing to him a psalm of praise."),
            new Verse("Praise", "Isaiah 25:1", "Lord, you are my God; I will exalt you and praise your name, for in perfect faithfulness you have done wonderful things, things planned long ago."),
            new Verse("Praise", "Psalm 103:1–2", "Praise the Lord, my soul; all my inmost being, praise his holy name. Praise the Lord, my soul, and forget not all his benefits."),
            new Verse("Praise", "Revelation 5:12", "In a loud voice they were saying: ‘Worthy is the Lamb, who was slain, to receive power and wealth and wisdom and strength and honor and glory and praise!’"),
            new Verse("Praise", "Psalm 145:3", "Great is the Lord and most worthy of praise; his greatness no one can fathom."),
            new Verse("Praise", "Ephesians 1:3", "Praise be to the God and Father of our Lord Jesus Christ, who has blessed us in the heavenly realms with every spiritual blessing in Christ."),
        };

        public Result Execute(ExternalCommandData commandData, ref string message, Autodesk.Revit.DB.ElementSet elements)
        {
            try
            {
                var v = Verses[_rng.Next(Verses.Count)];

                var sb = new StringBuilder();
                sb.AppendLine(v.Text);
                sb.AppendLine();
                sb.Append("— ").Append(v.Ref);

                var td = new TaskDialog("Inspirational Verse")
                {
                    MainInstruction = v.Category, // show the category
                    MainContent = sb.ToString(),
                    CommonButtons = TaskDialogCommonButtons.Ok
                };
                td.Show();

                // Clipboard copy intentionally removed to avoid extra references.

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"RandomVerseCommand failed: {ex.Message}";
                return Result.Failed;
            }
        }
    }
}
